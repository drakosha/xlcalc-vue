// Excel's calculation options over a session that calculates every edit (xlcalc records the option and does not apply
// it): in Manual the source keeps showing formulas' values as of the last calculation until Calculate Now or Calculate
// Sheet, and the workbook keeps the option.
import type { AsyncSession, BookInput, CellChange, CellValue } from 'xlcalc'
import { nameKey, parseCell } from 'xlcalc/a1'
import type { CalculationMode, CellUpdate, SheetCalculation, SheetWindow } from '../contract/source'

/** A cell's display as the sheet shows it. */
interface Shown {
  text: string
  value: CellValue
  color?: string
}

/** How many rectangles of formulas one `getCells` call reads while the values shown are captured. */
const RECTS_PER_READ = 2000

const shownOf = (c: { text?: string; value: CellValue; color?: string } | undefined): Shown => ({ text: c?.text ?? '', value: c ? c.value : null, ...(c?.color ? { color: c.color } : {}) })

/**
 * The values shown in Manual. `known` is captured from the session before the first change in Manual: every formula's
 * cell and the cells its array fills, as the sheet shows them; edited cells update it. A change of such a cell's value
 * that is not its own edit (a recalculation) is held: the sheet keeps the known value until a calculation shows the
 * session's. A cell outside `known` (a constant, an empty cell an array spills into anew) shows its change.
 */
export function manualCalculation(input: BookInput, emit: (calculation: SheetCalculation) => void) {
  let mode: CalculationMode = input.calculation ?? 'automatic'
  /** The values shown of formula cells by sheet key and address; null until the first change in Manual. */
  let known: Map<string, Map<string, Shown>> | null = null
  /** The cells whose session value is newer than the value shown, by sheet key: what a calculation updates. */
  const held = new Map<string, Map<string, { sheet: string; shown: Shown }>>()
  let told: SheetCalculation = { mode, pending: false }

  const manual = () => mode === 'manual'
  const pending = () => held.size > 0
  const state = (): SheetCalculation => ({ mode, pending: pending() })
  /** Tells subscribers when the option or whether cells wait changed. */
  function report(): void {
    const now = state()
    if (now.mode === told.mode && now.pending === told.pending) return
    told = now
    emit(now)
  }

  /** Captures the values shown before the first change in Manual: the session's, which nothing holds yet. */
  async function capture(s: AsyncSession): Promise<void> {
    // xlcalc before 0.66 has no Go To Special to find the formulas: the sheet shows calculated values, as in Automatic
    if (!manual() || known || typeof s.getSpecialCells !== 'function') return
    const next = new Map<string, Map<string, Shown>>()
    for (const { name } of (await s.getLayout()).sheets) {
      const cells = new Map<string, Shown>()
      const rects = await s.getSpecialCells(name, 'A1', { type: 'formulas' })
      const spills: string[] = []
      for (let i = 0; i < rects.length; i += RECTS_PER_READ) {
        for (const [addr, c] of Object.entries(await s.getCells(name, rects.slice(i, i + RECTS_PER_READ)))) {
          cells.set(addr, shownOf(c))
          if (c.spill && c.spill !== addr) spills.push(c.spill)
        }
      }
      // the cells an array fills past its formula are neither formulas nor constants
      for (let i = 0; i < spills.length; i += RECTS_PER_READ) {
        for (const [addr, c] of Object.entries(await s.getCells(name, spills.slice(i, i + RECTS_PER_READ)))) if (!cells.has(addr)) cells.set(addr, shownOf(c))
      }
      if (cells.size) next.set(nameKey(name), cells)
    }
    known = next
  }

  /** In Manual, a recalculated value of a known cell is held, and the sheet keeps showing the known one: true when held. */
  function hold(ch: CellChange): boolean {
    if (!manual() || !known || 'formula' in ch) return false
    const sk = nameKey(ch.sheet)
    const was = known.get(sk)?.get(ch.cell)
    if (!was) return false
    let mine = held.get(sk)
    if (!mine) held.set(sk, (mine = new Map()))
    if (!mine.has(ch.cell)) mine.set(ch.cell, { sheet: ch.sheet, shown: was })
    return true
  }
  /** Cells the change wrote or showed: what the sheet shows of them now is known and no longer waits. */
  function showing(updates: Iterable<CellUpdate>): void {
    if (!known) return
    for (const u of updates) {
      const sk = nameKey(u.sheet)
      let cells = known.get(sk)
      if (!cells) known.set(sk, (cells = new Map()))
      cells.set(u.cell, shownOf(u))
      const mine = held.get(sk)
      if (mine?.delete(u.cell) && !mine.size) held.delete(sk)
    }
  }
  /** Rows and columns moved (a structure operation or its undo): the values known by address no longer hold, so the sheet shows the session's. */
  function forget(): void {
    known = null
    held.clear()
  }

  /** The window with the values held on its sheet put back over the session's. */
  function overlay(sheet: string, win: SheetWindow, rows: number[]): SheetWindow {
    const mine = held.get(nameKey(sheet))
    if (!mine) return win
    const at = new Map(rows.map((r, i) => [r, i]))
    const width = win.text[0]?.length ?? 0
    for (const [addr, { shown }] of mine) {
      const p = parseCell(addr)
      const i = p ? at.get(p.row) : undefined
      if (!p || i === undefined || p.col < win.col || p.col >= win.col + width) continue
      win.text[i]![p.col - win.col] = shown.text
      win.values[i]![p.col - win.col] = shown.value
      if (shown.color) win.colors[addr] = shown.color
      else delete win.colors[addr]
    }
    return win
  }

  /** Calculate Now, or Calculate Sheet with `sheet`: the held cells (of that sheet) show the session's values; the caller sends them, then `report`. */
  async function calculate(s: AsyncSession, sheet?: string): Promise<CellUpdate[]> {
    const out: CellUpdate[] = []
    for (const [sk, mine] of [...held]) {
      if (sheet !== undefined && sk !== nameKey(sheet)) continue
      const name = mine.values().next().value!.sheet
      const got = await s.getCells(name, [...mine.keys()])
      for (const addr of mine.keys()) out.push({ sheet: name, cell: addr, ...shownOf(got[addr]) })
      held.delete(sk)
    }
    showing(out)
    return out
  }

  /** Sets the option: leaving Manual calculates what waits, as Excel does; the values known are dropped. */
  async function setMode(s: AsyncSession, next: CalculationMode): Promise<CellUpdate[]> {
    mode = next
    const out = next === 'manual' ? [] : await calculate(s)
    if (next !== 'manual') known = null
    return out
  }

  /** The workbook with the option as set here: a session keeps the option it was opened with. */
  function book<B extends BookInput>(b: B): B {
    const { calculation: _, ...rest } = b
    return (mode === 'automatic' ? rest : { ...rest, calculation: mode }) as B
  }

  return { state, report, capture, hold, showing, forget, overlay, calculate, setMode, book, manual }
}

export type ManualCalculation = ReturnType<typeof manualCalculation>
