// Excel's table commands over the source's `table`: Create Table (Insert Table, Format as Table, Ctrl+T and Ctrl+L) with its
// dialog, which guesses the range from the data around the active cell, and Convert to Range and Delete Table on the table at the active cell.
import { computed, ref } from 'vue'
import { cellAddress, colLetters, parseRange, rangeAddress, rectContains, sameName } from 'xlcalc/a1'
import type { CellPosition, CellRect } from 'xlcalc/contract'
import type { EditOutcome } from '../contract/source'
import type { TableOp } from '../contract/workbook'
import { cellRect } from '../rect'
import type { Commands } from './commands'
import { onMac, shortcutLetter } from './keys'
import { readRange, type ReadCell } from './range-read'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'
import type { Windows } from './windows'

/** The data read around the active cell at first to find its current region, widened while the region reaches the edge read. */
const REGION_ROWS = 200
const REGION_COLS = 30

/** The Create Table dialog: the range as typed (`=$A$1:$C$5`), My table has headers, and the reason the range was refused; or Convert to Range's question. */
export type TableDialog =
  | { kind: 'create'; sheet: string; draft: string; headers: boolean; error: string }
  | { kind: 'convert'; sheet: string; table: string }

type AddTable = Extract<TableOp, { op: 'addTable' }>

/** A range as Excel's Create Table box writes it: absolute, after `=` (`=$A$1:$C$5`). */
export const absoluteRange = (r: CellRect): string => {
  const a = `$${colLetters(r.left)}$${r.top}`
  return r.top === r.bottom && r.left === r.right ? `=${a}` : `=${a}:$${colLetters(r.right)}$${r.bottom}`
}

/**
 * Excel's current region around a cell (`Range.CurrentRegion`): the range grown, a row or column at a time, while the row or
 * column just outside it, corners included, has a non-empty cell. `filled` answers within `area`; cells outside it count as empty.
 */
export function currentRegion(at: CellPosition, area: CellRect, filled: (row: number, col: number) => boolean): CellRect {
  const r = cellRect(at.row, at.col)
  const any = (top: number, left: number, bottom: number, right: number): boolean => {
    for (let row = Math.max(top, area.top); row <= Math.min(bottom, area.bottom); row++)
      for (let col = Math.max(left, area.left); col <= Math.min(right, area.right); col++) if (filled(row, col)) return true
    return false
  }
  for (let grew = true; grew; ) {
    grew = false
    if (r.top > area.top && any(r.top - 1, r.left - 1, r.top - 1, r.right + 1)) (r.top--, (grew = true))
    if (r.bottom < area.bottom && any(r.bottom + 1, r.left - 1, r.bottom + 1, r.right + 1)) (r.bottom++, (grew = true))
    if (r.left > area.left && any(r.top - 1, r.left - 1, r.bottom + 1, r.left - 1)) (r.left--, (grew = true))
    if (r.right < area.right && any(r.top - 1, r.right + 1, r.bottom + 1, r.right + 1)) (r.right++, (grew = true))
  }
  return r
}

/**
 * Column names from the header cells' text, as Excel names them: an empty header `Column1` (the first number no other column
 * uses), a name another column to the left already has with the next free number (`Qty`, `Qty2`), compared case-insensitively.
 */
export function columnNames(headers: string[], prefix: string): string[] {
  const key = (x: string) => x.toLocaleLowerCase()
  const used = new Set<string>()
  for (const h of headers) if (h) used.add(key(h))
  const taken = new Set<string>()
  return headers.map((h) => {
    let name = h
    if (!h || taken.has(key(h))) {
      const base = h || prefix
      let n = h ? 2 : 1
      while (used.has(key(`${base}${n}`)) || taken.has(key(`${base}${n}`))) n++
      name = `${base}${n}`
    }
    taken.add(key(name))
    used.add(key(name))
    return name
  })
}

/** A header cell's text: a text value as it is, a number or a boolean as displayed, as Excel names a column by it whatever the width (`####` in a narrow column gives way to the edit text). */
function headerText(c: ReadCell | undefined): string {
  if (!c) return ''
  if (typeof c.value === 'string') return c.value
  return /^#+$/.test(c.text) ? (c.editText ?? String(c.value ?? '')) : c.text
}

export function useTables(s: SheetState, w: Windows, sel: Selection, cmd: Commands) {
  const read = useSheetReads(s)
  const dialog = ref<TableDialog | null>(null)
  const m = () => s.messages.value
  /** Create Table is offered: the source has `table` and the sheet is editable. */
  const available = computed(() => !!s.sheet.value && cmd.can('table'))
  /** My table has headers can be cleared: the header row Excel then inserts above the data needs `structure`. */
  const canInsertHeader = computed(() => cmd.can('structure'))
  /** The table at the active cell, from the loaded windows. */
  const activeTable = computed(() => {
    const a = s.selected.value
    return a ? (w.tables.value.find((t) => rectContains(t.rect, cellRect(a.row, a.col))) ?? null) : null
  })

  /** The range Excel's dialog proposes: a selection of several cells within the data, otherwise the active cell's current region. */
  async function guess(signal: AbortSignal): Promise<{ rect: CellRect; cells: Map<string, ReadCell> } | null> {
    const a = s.selected.value!
    const sheet = s.sheetName.value
    const r = sel.fullRect.value
    const data = { top: 1, left: 1, bottom: Math.max(1, s.dataRows.value), right: Math.max(1, s.dataCols.value) }
    if (r && (r.top !== r.bottom || r.left !== r.right)) {
      const rect = { top: r.top, left: r.left, bottom: Math.max(r.top, Math.min(r.bottom, data.bottom)), right: Math.max(r.left, Math.min(r.right, data.right)) }
      const cells = await readRange(s.source(), sheet, { ...rect, bottom: Math.min(rect.bottom, rect.top + 1) }, signal)
      return cells && { rect, cells }
    }
    for (let rows = REGION_ROWS, cols = REGION_COLS; ; rows *= 4, cols *= 4) {
      const area = { top: Math.max(1, a.row - rows), left: Math.max(1, a.col - cols), bottom: Math.min(data.bottom, a.row + rows), right: Math.min(data.right, a.col + cols) }
      if (area.top > area.bottom || area.left > area.right || a.row > data.bottom || a.col > data.right) return { rect: cellRect(a.row, a.col), cells: new Map() }
      const cells = await readRange(s.source(), sheet, area, signal)
      if (!cells) return null
      const rect = currentRegion(a, area, (row, col) => cells.has(cellAddress(row, col)))
      const edge = (rect.top === area.top && area.top > 1) || (rect.left === area.left && area.left > 1) || (rect.bottom === area.bottom && area.bottom < data.bottom) || (rect.right === area.right && area.right < data.right)
      if (!edge) return { rect, cells }
    }
  }

  /**
   * Excel's Create Table (Ctrl+T, Ctrl+L): the dialog with the selection, or the active cell's current region, and My table has
   * headers checked when its first row holds text in every column above more rows. False when the sheet has no table commands.
   */
  async function openCreate(): Promise<boolean> {
    if (!available.value || !s.selected.value) return false
    const sheet = s.sheetName.value
    const got = await guess(read())
    if (!got || sheet !== s.sheetName.value) return false
    const { rect, cells } = got
    const first = Array.from({ length: rect.right - rect.left + 1 }, (_, i) => cells.get(cellAddress(rect.top, rect.left + i)))
    // without `structure` there is no header row to insert above the data: the box stays checked
    const headers = !canInsertHeader.value || (rect.bottom > rect.top && first.every((c) => !!c && typeof c.value === 'string' && c.value !== '' && c.formula === undefined))
    dialog.value = { kind: 'create', sheet, draft: absoluteRange(rect), headers, error: '' }
    return true
  }

  /** The range typed in the dialog (`=$A$1:$C$5`, `A1:C5`, with the sheet's name or without), null when Excel would not take it. */
  function typedRange(draft: string, sheet: string): CellRect | null {
    let text = draft.trim().replace(/^=/, '')
    const bang = text.lastIndexOf('!')
    if (bang >= 0) {
      const name = text.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'")
      if (!sameName(name, sheet)) return null
      text = text.slice(bang + 1)
    }
    const r = parseRange(text)
    return r?.kind === 'cells' ? r : null
  }

  /** Every table's name in the workbook: each sheet's first window lists its tables. */
  async function tableNames(signal: AbortSignal): Promise<string[]> {
    const source = s.source()
    const lists = await Promise.all(
      s.sheets.value.map((x) =>
        source.window({ sheet: x.name, row: 1, rows: 1, col: 1, cols: 1 }, signal).then(
          (got) => ('pending' in got ? [] : (got.tables ?? []).map((t) => t.name)),
          () => [],
        ),
      ),
    )
    return lists.flat()
  }
  /** Excel's name for a new table, `Table1`, `Table2`: the first number from the count of tables on that no table and no defined name has. */
  async function nextName(signal: AbortSignal): Promise<string> {
    const prefix = m().tableNamePrefix
    const names = s.sourceMethod('names')
    const [tables, defined] = await Promise.all([tableNames(signal), names ? names(signal).catch(() => []) : []])
    const used = new Set([...tables, ...defined.map((n) => n.name)].map((x) => x.toLocaleLowerCase()))
    let n = tables.length + 1
    while (used.has(`${prefix}${n}`.toLocaleLowerCase())) n++
    return `${prefix}${n}`
  }

  /** The table to add. Its style (Excel's gallery, xlcalc's `TableInput.style`) is chosen here once the sheet offers one; without it the source writes the default. */
  const tableOf = (sheet: string, name: string, range: CellRect, columns: string[]): AddTable => ({
    op: 'addTable',
    sheet,
    table: { name, range: rangeAddress(range), headers: true, columns: columns.map((c) => ({ name: c })) },
  })

  /**
   * OK in Create Table: the table over the range, its columns named after the header cells; without My table has headers,
   * Excel's header row of `Column1`, `Column2`… above the data, which first moves the cells from the range's first row down
   * by one row in its columns (a `moveRange` through `structure`, then the table: two undo steps).
   */
  async function create(): Promise<EditOutcome | null> {
    const d = dialog.value
    if (d?.kind !== 'create') return null
    const r = typedRange(d.draft, d.sheet)
    if (!r || d.sheet !== s.sheetName.value) {
      d.error = m().tableRangeBad
      return null
    }
    const signal = read()
    const sheet = d.sheet
    const width = r.right - r.left + 1
    const headers = d.headers || !canInsertHeader.value
    let names: string[]
    if (headers) {
      const cells = await readRange(s.source(), sheet, { ...r, bottom: r.top }, signal)
      if (!cells) return null
      names = columnNames(Array.from({ length: width }, (_, i) => headerText(cells.get(cellAddress(r.top, r.left + i)))), m().tableColumnPrefix)
    } else names = columnNames(new Array<string>(width).fill(''), m().tableColumnPrefix)
    const name = await nextName(signal)
    if (signal.aborted || sheet !== s.sheetName.value) return null
    dialog.value = null
    s.focusGrid()
    let range = r
    if (!headers) {
      const below = Math.max(r.bottom, s.dataRows.value)
      const moved = await cmd.structure({ op: 'moveRange', from: { sheet, range: rangeAddress({ ...r, bottom: below }) }, to: { sheet, cell: cellAddress(r.top + 1, r.left) } })
      if (!moved?.ok) return moved
      range = { ...r, bottom: r.bottom + 1 }
    }
    return cmd.table(tableOf(sheet, name, range, names))
  }

  /** Excel's Convert to Range on the table at the active cell, after its question; null without one or on No. */
  let answer: ((yes: boolean) => void) | null = null
  async function convert(): Promise<EditOutcome | null> {
    const t = activeTable.value
    if (!t || !available.value) return null
    answer?.(false)
    dialog.value = { kind: 'convert', sheet: s.sheetName.value, table: t.name }
    const yes = await new Promise<boolean>((resolve) => (answer = resolve))
    return yes ? cmd.table({ op: 'convertTable', table: t.name }) : null
  }
  /** Delete Table: the table at the active cell with its cells; null without one. */
  async function remove(): Promise<EditOutcome | null> {
    const t = activeTable.value
    return t && available.value ? cmd.table({ op: 'deleteTable', table: t.name }) : null
  }
  /** Closes the dialog: Cancel in Create Table, Yes or No in Convert to Range. */
  function close(yes = false): void {
    dialog.value = null
    const done = answer
    answer = null
    done?.(yes)
    s.focusGrid()
  }

  /** Ctrl+T and Ctrl+L, Excel's Create Table (Control+T only on a Mac, where Control+L is Define Name and the browser keeps Cmd+T); true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (!e.ctrlKey || e.altKey || e.shiftKey || e.metaKey || !available.value) return false
    const letter = shortcutLetter(e)
    // Control+L on a Mac is Define Name, as in Excel for Mac
    if (letter !== 't' && (letter !== 'l' || onMac())) return false
    e.preventDefault()
    void openCreate()
    return true
  }

  return { available, canInsertHeader, activeTable, dialog, openCreate, create, convert, remove, close, key }
}

export type Tables = ReturnType<typeof useTables>
