// Excel's Go To Special (Home > Find & Select, the Go To dialog's Special…): the dialog's state and the selection it makes.
// Blanks, constants, formulas, visible cells and the last cell come from the source's `specialCells` (xlcalc's getSpecialCells);
// the current region is read from the windows, as Create Table reads it; conditional formats and data validation come from
// the sheet's rules (`rules`), by Excel's rules for All and Same as measured in Excel for Mac. Several rectangles make a
// selection of several areas.
import { computed, ref } from 'vue'
import { parseRange, rangeAddress } from 'xlcalc/a1'
import type { CellPosition, CellRect, SpecialValue } from 'xlcalc/contract'
import type { SheetSpecialCells } from '../contract/source'
import type { SheetRules } from '../contract/workbook'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'
import { readRegion } from './tables'
import type { AlertAnswer, SheetAlert } from './validation'

/** The options of Excel's Go To Special the component offers. */
export type SpecialKind = 'constants' | 'formulas' | 'blanks' | 'currentRegion' | 'lastCell' | 'visible' | 'conditionalFormats' | 'validation'

/** The dialog's choice: an option, the kinds of values for Constants and Formulas, Same (rather than All) for the rules. */
export interface SpecialChoice {
  kind: SpecialKind
  values: SpecialValue[]
  same: boolean
}

export const SPECIAL_VALUES: SpecialValue[] = ['numbers', 'text', 'logicals', 'errors']

export interface SpecialDeps {
  ask: (alert: SheetAlert) => Promise<AlertAnswer>
  notify: (text: string) => void
}

/** The rectangles of ranges written as xlsx `sqref` (`B2:B9 D2`, `C:C`). */
const rectsOf = (ranges: string[]): CellRect[] => ranges.flatMap((r) => r.split(/\s+/)).flatMap((r) => parseRange(r) ?? [])
const holds = (r: CellRect, p: CellPosition) => p.row >= r.top && p.row <= r.bottom && p.col >= r.left && p.col <= r.right
/** The cells two ranges share; null when they do not meet. */
function meet(a: CellRect, b: CellRect): CellRect | null {
  const r = { top: Math.max(a.top, b.top), left: Math.max(a.left, b.left), bottom: Math.min(a.bottom, b.bottom), right: Math.min(a.right, b.right) }
  return r.top <= r.bottom && r.left <= r.right ? r : null
}
/** A value as text with its keys sorted, to compare two rules field by field. */
function canonical(x: unknown): string {
  if (Array.isArray(x)) return `[${x.map(canonical).join(',')}]`
  if (x && typeof x === 'object')
    return `{${Object.keys(x)
      .sort()
      .map((k) => `${JSON.stringify(k)}:${canonical((x as Record<string, unknown>)[k])}`)
      .join(',')}}`
  return JSON.stringify(x)
}

/**
 * The cells of the sheet's rules that Go To Special selects, as Excel for Mac selected them: All takes every rule's cells,
 * within the searched areas unless they are a single cell (then the whole sheet); Same ignores the areas and takes, for
 * conditional formats, the cells of the first rule (by priority) that covers the active cell, and for data validation the cells
 * of every rule with the same settings as the active cell's (Excel joins two such rules).
 */
export function ruleCells(rules: SheetRules, kind: 'conditionalFormats' | 'validation', same: boolean, at: CellPosition, areas: CellRect[] | null): CellRect[] {
  const list = kind === 'conditionalFormats' ? rules.conditionalFormats.map((r) => ({ rule: r as object, rects: rectsOf(r.ranges) })) : rules.validations.map((r) => ({ rule: r as object, rects: rectsOf([r.range]) }))
  if (!same) return areas ? list.flatMap((x) => x.rects.flatMap((r) => areas.flatMap((a) => meet(r, a) ?? []))) : list.flatMap((x) => x.rects)
  const mine = list.find((x) => x.rects.some((r) => holds(r, at)))
  if (!mine) return []
  if (kind === 'conditionalFormats') return mine.rects
  const settings = (rule: object) => canonical({ ...rule, range: null })
  const key = settings(mine.rule)
  return list.filter((x) => settings(x.rule) === key).flatMap((x) => x.rects)
}

export function useSpecial(s: SheetState, sel: Selection, deps: SpecialDeps) {
  const read = useSheetReads(s)
  const open = ref(false)
  /** The last choice, which the dialog opens with, as Excel keeps it: Constants of every kind at first (Excel's Notes is not offered). */
  const last = ref<SpecialChoice>({ kind: 'constants', values: [...SPECIAL_VALUES], same: false })
  /** Go To Special is offered: a sheet with the source's `specialCells`, and no display filter, whose rows on screen are not the sheet's. */
  const available = computed(() => !!s.sheet.value && !s.visible.value && typeof s.source().specialCells === 'function')
  /** Conditional formats and Data validation need the sheet's rules. */
  const hasRules = computed(() => typeof s.source().rules === 'function')

  function openSpecial(): boolean {
    if (!available.value) return false
    open.value = true
    return true
  }
  function close(): void {
    open.value = false
    s.focusGrid()
  }

  /** The rectangles the choice finds over the selection: each of its areas, or the used range for a single cell. */
  async function find(choice: SpecialChoice, signal: AbortSignal): Promise<CellRect[] | null> {
    const source = s.source()
    const sheet = s.sheetName.value
    const at = s.selected.value!
    const full = sel.fullRect.value!
    const one = !s.areas.value && full.top === full.bottom && full.left === full.right
    const areas = s.areas.value ?? [full]
    const rects = (list: string[]) => list.flatMap((r) => parseRange(r) ?? [])
    if (choice.kind === 'currentRegion') {
      const got = await readRegion(s, at, signal)
      return got && [got.rect]
    }
    if (choice.kind === 'conditionalFormats' || choice.kind === 'validation') return ruleCells(await source.rules!(sheet, signal), choice.kind, choice.same, at, one ? null : areas)
    const kind: SheetSpecialCells =
      choice.kind === 'constants' || choice.kind === 'formulas' ? { type: choice.kind, ...(choice.values.length < SPECIAL_VALUES.length ? { values: choice.values } : {}) } : choice.kind === 'lastCell' ? { type: 'lastCell' } : { type: choice.kind }
    if (choice.kind === 'lastCell') return rects(await source.specialCells!(sheet, 'A1', kind, signal))
    const found = await Promise.all(areas.map((a) => source.specialCells!(sheet, rangeAddress(a), kind, signal)))
    return found.flatMap(rects)
  }

  /**
   * The dialog's OK: it closes and the cells found are selected, the active cell kept when one of them holds it; Excel's "No
   * cells were found." when there are none, the selection unchanged. A failed read goes to the message line.
   */
  async function run(choice: SpecialChoice): Promise<void> {
    last.value = choice
    close()
    if (!s.selected.value || !sel.fullRect.value) return
    const outdated = s.outdated()
    let got: CellRect[] | null
    try {
      got = await find(choice, read())
    } catch (e) {
      if (!outdated()) deps.notify(e instanceof Error ? e.message : String(e))
      return
    }
    if (outdated() || got === null) return
    if (!got.length) {
      await deps.ask({ severity: 'notice', message: s.messages.value.specialNone })
      return
    }
    sel.selectAreas(got, s.selected.value)
  }

  return { open, last, available, hasRules, openSpecial, close, run }
}

export type Special = ReturnType<typeof useSpecial>
