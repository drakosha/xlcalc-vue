// Excel's calculation options over a session that calculates every edit (xlcalc records the option and does not apply
// it): in Manual the source keeps showing formulas' values as of the last calculation until Calculate Now or Calculate
// Sheet, and the workbook keeps the option.
import type { AddressMove, AsyncSession, BookInput, CellChange, CellValue, CellView } from 'xlcalc'
import { nameKey, parseCell } from 'xlcalc/a1'
import { cellKey } from './outcome'
import type { CalculationMode, CellUpdate, SheetCalculation, SheetWindow } from '../contract/source'

/** A cell's display as the sheet shows it; `of` is the formula whose array fills the cell. */
interface Shown {
  text: string
  value: CellValue
  color?: string
  of?: string
}

/** How many rectangles of formulas one `getCells` call reads while the values shown are captured. */
const RECTS_PER_READ = 2000

/** Whether two values are the same, an error by its code. */
const sameValue = (a: CellValue, b: CellValue): boolean => a === b || (typeof a === 'object' && a !== null && typeof b === 'object' && b !== null && a.error === b.error)

const shownOf = (c: { text?: string; value: CellValue; color?: string; spillOf?: string } | undefined): Shown => ({ text: c?.text ?? '', value: c ? c.value : null, ...(c?.color ? { color: c.color } : {}), ...(c?.spillOf ? { of: c.spillOf } : {}) })

/** A cell as the session reads it, as a change that lists it: its display, whether it has a formula, the array that fills it. */
const changeOf = (sheet: string, cell: string, c: CellView | undefined): CellChange => ({
  sheet,
  cell,
  text: c?.text ?? '',
  value: c ? c.value : null,
  ...(c?.color ? { color: c.color } : {}),
  formula: c?.formula ?? null,
  ...(c?.spillOf ? { spillOf: c.spillOf } : {}),
})

/**
 * The values shown in Manual. `known` is captured from the session before the first change in Manual: every formula's
 * cell and the cells its array fills, as the sheet shows them; the cells that changes show update it, a formula entered
 * joining it and a cell left without one leaving it. A change of such a cell's value that is not its own edit (a
 * recalculation) is held: the sheet keeps the known value until a calculation shows the session's. A change of how such a
 * cell looks with its value as known (a number format, a column too narrow, a note) shows, as does a cell of the array
 * of a formula entered, which Excel calculates with its cell. A cell outside `known` (a constant, an empty cell an array
 * spills into anew) shows its change.
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

  /** In Manual, a recalculated value of a known cell is held, and the sheet keeps showing the known one: true when held. `entered`: the cells entered, by `cellKey`. */
  function hold(ch: CellChange, entered: Set<string>): boolean {
    if (!manual() || !known || 'formula' in ch) return false
    const sk = nameKey(ch.sheet)
    const was = known.get(sk)?.get(ch.cell)
    if (!was || sameValue(was.value, ch.value)) return false
    // a cell of the array of a formula entered, filled now or left empty
    if (was.of && entered.has(cellKey(ch.sheet, was.of))) return false
    let mine = held.get(sk)
    if (!mine) held.set(sk, (mine = new Map()))
    if (!mine.has(ch.cell)) mine.set(ch.cell, { sheet: ch.sheet, shown: was })
    return true
  }
  /**
   * Cells the change wrote or showed, with the session's `changes` that listed them: they no longer wait, and what the
   * sheet shows of a formula's cell is known. An edited cell's formula, or a change's, tells whether the cell has one;
   * a cell an array fills is known; a cell with neither keeps what it was.
   */
  function showing(updates: Iterable<CellUpdate>, changes: CellChange[] = []): void {
    if (!known) return
    const listed = new Map(changes.map((c) => [cellKey(c.sheet, c.cell), c]))
    for (const u of updates) {
      const sk = nameKey(u.sheet)
      const mine = held.get(sk)
      if (mine?.delete(u.cell) && !mine.size) held.delete(sk)
      const ch = listed.get(cellKey(u.sheet, u.cell))
      let cells = known.get(sk)
      const was = cells?.get(u.cell)
      // a cell no change tells about stays a formula's, and an array's cell without its array is left empty
      const kept = !!ch?.spillOf || (u.formula !== undefined ? u.formula !== null : ch && 'formula' in ch ? ch.formula !== null : !!was && !was.of)
      if (!kept) {
        cells?.delete(u.cell)
        continue
      }
      if (!cells) known.set(sk, (cells = new Map()))
      cells.set(u.cell, { ...shownOf(u), ...(ch?.spillOf ? { of: ch.spillOf } : {}) })
    }
  }
  /**
   * The waiting cells of the arrays of formulas `entered` (an edit's cells), as changes: entering a formula again
   * calculates its whole array, the cells it no longer fills too, though the session's values of them did not change.
   */
  async function spillsOf(s: AsyncSession, entered: Iterable<CellUpdate>): Promise<CellChange[]> {
    if (!held.size) return []
    const keys = new Set([...entered].map((u) => cellKey(u.sheet, u.cell)))
    const out: CellChange[] = []
    for (const mine of held.values()) {
      const sheet = mine.values().next().value!.sheet
      const cells = [...mine].filter(([, { shown }]) => shown.of && keys.has(cellKey(sheet, shown.of))).map(([addr]) => addr)
      if (!cells.length) continue
      const got = await s.getCells(sheet, cells)
      for (const addr of cells) out.push(changeOf(sheet, addr, got[addr]))
    }
    return out
  }

  /** Values known by address that no longer hold (a table command): the sheet shows the session's. */
  function forget(): void {
    known = null
    held.clear()
  }
  /**
   * How a structure operation or its undo moved addresses (`EditResult.moved`): rows, columns or a range moved leave the
   * values known by address behind, so the sheet shows the session's (`forget`); a renamed sheet keeps its own under its
   * new name and a deleted one takes them with it. A paste, a fill, a merge, a sheet inserted, moved or hidden move nothing.
   */
  function moved(moves: AddressMove[] = []): void {
    if (moves.some((m) => m.kind === 'lines' || m.kind === 'range')) return forget()
    for (const m of moves) {
      if (m.kind !== 'deleteSheet' && m.kind !== 'renameSheet') continue
      const from = nameKey(m.sheet)
      const cells = known?.get(from)
      const mine = held.get(from)
      known?.delete(from)
      held.delete(from)
      if (m.kind === 'deleteSheet') continue
      if (cells) known!.set(nameKey(m.name), cells)
      if (mine) held.set(nameKey(m.name), new Map([...mine].map(([addr, h]) => [addr, { ...h, sheet: m.name }])))
    }
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
    const changes: CellChange[] = []
    for (const [sk, mine] of [...held]) {
      if (sheet !== undefined && sk !== nameKey(sheet)) continue
      const name = mine.values().next().value!.sheet
      const got = await s.getCells(name, [...mine.keys()])
      for (const addr of mine.keys()) {
        const ch = changeOf(name, addr, got[addr])
        changes.push(ch)
        out.push({ sheet: name, cell: addr, text: ch.text ?? '', value: ch.value, ...(ch.color ? { color: ch.color } : {}) })
      }
      held.delete(sk)
    }
    showing(out, changes)
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

  return { state, report, capture, hold, showing, spillsOf, forget, moved, overlay, calculate, setMode, book, manual }
}

export type ManualCalculation = ReturnType<typeof manualCalculation>
