// A timeline's periods and its selection label: the years, quarters, months or days between its bounds (whole years, as
// xlcalc bounds it), the dates a run of periods chooses, and the label Excel shows over a selection, in the messages' words.
import type { TimelineLevel } from 'xlcalc/contract'
import type { SlicerMessages } from './contract/messages'

/** A period of a timeline: its first and last day (ISO 8601 without a zone, as xlcalc's dates), its tile, its year's or month's group. */
export interface TimelinePeriod {
  start: string
  end: string
  label: string
  group: string
}

const pad = (n: number, w = 2) => String(n).padStart(w, '0')
/** A day as xlcalc writes a date. */
const iso = (d: Date) => `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}T00:00:00`
/** A date of xlcalc's as a day; null for one that is not. */
function day(text: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(text)
  return m ? new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]))) : null
}
const MONTHS_IN: Record<Exclude<TimelineLevel, 'days'>, number> = { years: 12, quarters: 3, months: 1 }

/** The start of the period of `level` that holds a day. */
function periodStart(d: Date, level: TimelineLevel): Date {
  if (level === 'days') return d
  const m = d.getUTCMonth()
  return new Date(Date.UTC(d.getUTCFullYear(), m - (m % MONTHS_IN[level]), 1))
}
/** The start of the next period. */
function nextStart(d: Date, level: TimelineLevel): Date {
  if (level === 'days') return new Date(d.getTime() + 86_400_000)
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + MONTHS_IN[level], 1))
}

/** A period's tile: the year, the quarter, the month's short name, the day of the month. */
function tile(d: Date, level: TimelineLevel, m: SlicerMessages): string {
  if (level === 'years') return String(d.getUTCFullYear())
  if (level === 'quarters') return m.timelineQuarter(Math.floor(d.getUTCMonth() / 3) + 1)
  if (level === 'months') return m.timelineMonths[d.getUTCMonth()] ?? ''
  return String(d.getUTCDate())
}

/** The periods of `level` from the bounds' start to their end (exclusive); a timeline without bounds has none. */
export function timelinePeriods(bounds: { start: string; end: string } | undefined, level: TimelineLevel, m: SlicerMessages): TimelinePeriod[] {
  const from = bounds && day(bounds.start)
  const to = bounds && day(bounds.end)
  if (!from || !to) return []
  const out: TimelinePeriod[] = []
  for (let d = periodStart(from, level); d < to && out.length < 40_000; ) {
    const next = nextStart(d, level)
    const group = level === 'years' ? '' : level === 'days' ? m.timelinePeriod(m.timelineMonths[d.getUTCMonth()] ?? '', d.getUTCFullYear()) : String(d.getUTCFullYear())
    out.push({ start: iso(d), end: iso(new Date(next.getTime() - 86_400_000)), label: tile(d, level, m), group })
    d = next
  }
  return out
}

/** Whether a period lies in the dates chosen, by whole days. */
export const inSelection = (p: TimelinePeriod, s: { start: string; end: string } | undefined): boolean => !!s && p.start.slice(0, 10) <= s.end.slice(0, 10) && p.end.slice(0, 10) >= s.start.slice(0, 10)

/**
 * The label over a timeline's selection, at the level it was chosen at: All Periods without one; one day; whole years; one
 * period or a run of periods within a year or across years, in the patterns of Excel's string table (not measured).
 */
export function selectionLabel(s: { start: string; end: string } | undefined, level: TimelineLevel, m: SlicerMessages): string {
  const a = s && day(s.start)
  const b = s && day(s.end)
  if (!a || !b) return m.timelineAllPeriods
  const dayText = (d: Date) => m.timelineDay(d.getUTCDate(), m.timelineMonths[d.getUTCMonth()] ?? '', d.getUTCFullYear())
  if (a.getTime() === b.getTime()) return dayText(a)
  if (level === 'days') return m.timelineYears(dayText(a), dayText(b))
  const ya = a.getUTCFullYear()
  const yb = b.getUTCFullYear()
  if (level === 'years') return ya === yb ? String(ya) : m.timelineYears(String(ya), String(yb))
  const ta = tile(a, level, m)
  const tb = tile(b, level, m)
  if (ya !== yb) return m.timelineAcross(ta, ya, tb, yb)
  return ta === tb ? m.timelinePeriod(ta, ya) : m.timelineInYear(ta, tb, ya)
}
