// A source over a book in memory, for tests and examples: values, shown text, formulas as strings and every display field; no engine.
// It has every optional method of 0.2.0 (search, filter, column values, totals), Replace All, and can play a book being recomputed.
import { cellAddress, parseCell, parseLines, parseRange, rectsMeet } from 'xlcalc/a1'
import type { CellStyle, CellValue, ConditionalAppearance } from 'xlcalc/contract'
import { textEditText } from './core/edit-text'
import type { CellEntry, CellUpdate, EditOutcome, LayoutChange, ReplaceOutcome, SheetInfo, SheetPending, SheetSource, SheetWindow, SourceEvent, WindowRequest } from './source'
import { columnList, findHits, rangeStats, replaceTexts, visibleList, type QuerySheet } from './memory-query'

/** A cell of the memory book; a bare value is a cell with that value. */
export interface MemoryCell {
  value?: CellValue
  /** Shown text; by default the value as plain text. */
  text?: string
  /** The formula as shown, without `=`; with no engine its value is whatever `value` says. */
  formula?: string
  /** A style of the book's `styles`, by index. */
  style?: number
  /** The colour of the number format's section. */
  color?: string
  link?: { href?: string; location?: string; tooltip?: string; target?: string }
  note?: { author?: string; text: string }
  /** An old-style data bar (`SheetWindow.bars`); `appearance` takes its place where both are given. */
  bar?: { ratio: number; color: string }
  /** What conditional formats give the cell, as xlcalc computes it (`ConditionalAppearance`); the book in memory keeps it as given. */
  appearance?: ConditionalAppearance
}

export interface MemorySheet extends Omit<SheetInfo, 'maxRow' | 'maxCol'> {
  /** Cells by A1 address. */
  cells: Record<string, MemoryCell | CellValue>
  /** Column widths by letters (`B`) or a span (`B:D`), in xlsx units; 0 hides. */
  widths?: Record<string, number>
  /** Row heights by row (`5`) or a span (`5:9`), in points; 0 hides. */
  heights?: Record<string, number>
  defaultRowHeight?: number
  merges?: string[]
  tables?: { name: string; range: string; headerRow: number }[]
  notice?: string
}

export interface MemoryBook {
  sheets: MemorySheet[]
  styles?: CellStyle[]
}

export interface MemorySourceOptions {
  /** Milliseconds every answer takes, to see the component wait; 0 by default. */
  delay?: number
  maxCells?: number
}

export interface MemorySource extends SheetSource {
  /** Answer the next edits with these refusals instead of applying them (tests and the example). */
  failNext(...outcomes: Extract<EditOutcome, { ok: false }>[]): void
  /** Send an event to the subscribers, as another tab would. */
  emit(event: SourceEvent): void
  /** Change cells as another tab would: applied, and sent to the subscribers as `changed`. */
  setCells(sheet: string, entries: Record<string, string | null>): CellUpdate[]
  /** Play a book being recomputed for `ms`: windows answer `pending` until then. */
  recompute(ms: number): void
}

const ERROR_TEXT: Record<string, string> = {
  DIV0: '#DIV/0!',
  NA: '#N/A',
  NAME: '#NAME?',
  NUM: '#NUM!',
  REF: '#REF!',
  VALUE: '#VALUE!',
  NULL: '#NULL!',
  CYCLE: '#CYCLE!',
  SPILL: '#SPILL!',
  CALC: '#CALC!',
  ERROR: '#ERROR!',
}

/** A value as plain text: numbers as JavaScript writes them, booleans and errors as Excel does. */
function plainText(value: CellValue | undefined): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'object') return ERROR_TEXT[value.error] ?? '#ERROR!'
  return String(value)
}

interface Store {
  sheet: MemorySheet
  cells: Map<string, MemoryCell>
  widths: Map<number, number>
  heights: Map<number, number>
  maxRow: number
  maxCol: number
}

/** Each column or row of a key as a book takes it (`B`, `B:D`; `5`, `5:9`); none for a key it does not take. */
function linesOf(key: string, axis: 'columns' | 'rows'): number[] {
  const span = parseLines(key, axis)
  const out: number[] = []
  if (span) for (let i = span.first; i <= span.last; i++) out.push(i)
  return out
}

const abortError = () => new DOMException('The request was aborted', 'AbortError')

/** A source over `book`; the book's objects are copied, edits change the copy. */
export function memorySource(book: MemoryBook, options: MemorySourceOptions = {}): MemorySource {
  const delay = options.delay ?? 0
  const styles = book.styles ?? []
  const stores = book.sheets.map((sheet): Store => {
    const cells = new Map<string, MemoryCell>()
    for (const [addr, raw] of Object.entries(sheet.cells)) {
      const cell = raw !== null && typeof raw === 'object' && !('error' in raw) ? { ...raw } : { value: raw as CellValue }
      cells.set(addr.toUpperCase(), cell)
    }
    const widths = new Map<number, number>()
    for (const [key, w] of Object.entries(sheet.widths ?? {})) for (const c of linesOf(key, 'columns')) widths.set(c, w)
    const heights = new Map<number, number>()
    for (const [key, h] of Object.entries(sheet.heights ?? {})) for (const r of linesOf(key, 'rows')) heights.set(r, h)
    const store: Store = { sheet, cells, widths, heights, maxRow: 0, maxCol: 0 }
    measure(store)
    return store
  })
  const listeners = new Set<(event: SourceEvent) => void>()
  const failures: Extract<EditOutcome, { ok: false }>[] = []
  /** A recomputation being played: when it started and when it ends. */
  let recomputing: { since: number; until: number } | null = null

  /** A sheet as the queries read it: shown text and values by row and column. */
  function query(store: Store): QuerySheet {
    return {
      maxRow: store.maxRow,
      maxCol: store.maxCol,
      textAt: (r, c) => {
        const cell = store.cells.get(cellAddress(r, c))
        return cell ? textOf(cell) : ''
      },
      valueAt: (r, c) => {
        const cell = store.cells.get(cellAddress(r, c))
        return cell?.value ?? (cell?.text ? cell.text : null)
      },
      editAt: (r, c) => {
        const cell = store.cells.get(cellAddress(r, c))
        return cell ? editTextOf(cell) : ''
      },
    }
  }
  function storeOrThrow(name: string): Store {
    const store = storeOf(name)
    if (!store) throw new Error(`No sheet ${name}`)
    return store
  }

  /** The sheet's extent: cells with anything in them, and merges. */
  function measure(store: Store): void {
    let maxRow = 0
    let maxCol = 0
    for (const addr of store.cells.keys()) {
      const at = parseCell(addr)
      if (!at) continue
      maxRow = Math.max(maxRow, at.row)
      maxCol = Math.max(maxCol, at.col)
    }
    for (const m of store.sheet.merges ?? []) {
      const r = parseRange(m)
      if (r) {
        maxRow = Math.max(maxRow, r.bottom)
        maxCol = Math.max(maxCol, r.right)
      }
    }
    store.maxRow = maxRow
    store.maxCol = maxCol
  }
  const storeOf = (name: string) => stores.find((x) => x.sheet.name === name)
  const textOf = (cell: MemoryCell) => cell.text ?? plainText(cell.value)
  /** What a cell opens with as `apply` reads it back: `=` and its formula, a text with Excel's apostrophe where it would read as a value, else the value as plain text. */
  function editTextOf(cell: MemoryCell): string {
    if (cell.formula !== undefined) return `=${cell.formula}`
    if (typeof cell.value === 'string') return textEditText(cell.value)
    return cell.value === undefined || cell.value === null ? (cell.text ?? '') : plainText(cell.value)
  }

  function wait<T>(make: () => T, signal?: AbortSignal): Promise<T> {
    return new Promise((resolve, reject) => {
      if (signal?.aborted) return reject(abortError())
      const done = () => {
        if (signal?.aborted) reject(abortError())
        else {
          try {
            resolve(make())
          } catch (e) {
            reject(e)
          }
        }
      }
      if (delay > 0) setTimeout(done, delay)
      else queueMicrotask(done)
    })
  }

  function info(store: Store): SheetInfo {
    const { name, freeze, hidden, locked, columns, rowNumbers, totals } = store.sheet
    return { name, maxRow: store.maxRow, maxCol: store.maxCol, freeze, hidden, locked, columns, rowNumbers, totals }
  }

  /** A rectangle; under a filter `row` is a display index and the rows come from the filter's list. */
  function windowOf(req: WindowRequest): SheetWindow | SheetPending {
    if (recomputing && Date.now() < recomputing.until) return { pending: { elapsedMs: Date.now() - recomputing.since } }
    const store = storeOrThrow(req.sheet)
    const list = visibleList(query(store), req.filter)
    const cols = Math.max(0, req.cols)
    const limit = list ? Math.max(0, list.length - req.row + 1) : Infinity
    const rows = Math.max(0, Math.min(req.rows, limit, Math.floor((options.maxCells ?? 20_000) / Math.max(1, cols))))
    const win: SheetWindow = {
      row: req.row,
      col: req.col,
      text: [],
      values: [],
      formulas: {},
      colors: {},
      styleIds: {},
      styles,
      widths: [],
      rowHeights: {},
      merges: [],
      defaultRowHeight: store.sheet.defaultRowHeight,
      tables: store.sheet.tables ?? [],
      bars: {},
      appearance: {},
      links: {},
      notes: {},
      ...(store.sheet.notice ? { notice: store.sheet.notice } : {}),
    }
    for (let c = req.col; c < req.col + cols; c++) win.widths.push(store.widths.get(c) ?? null)
    const sheetRows: number[] = []
    for (let d = req.row; d < req.row + rows; d++) {
      const r = list ? list[d - 1]! : d
      sheetRows.push(r)
      const text: string[] = []
      const values: CellValue[] = []
      const h = store.heights.get(r)
      if (h !== undefined) win.rowHeights[r] = h
      for (let c = req.col; c < req.col + cols; c++) {
        const addr = cellAddress(r, c)
        const cell = store.cells.get(addr)
        text.push(cell ? textOf(cell) : '')
        values.push(cell?.value ?? (cell?.text ? cell.text : null))
        if (!cell) continue
        if (cell.formula !== undefined) win.formulas[addr] = cell.formula
        if (cell.color) win.colors[addr] = cell.color
        if (cell.style !== undefined) win.styleIds[addr] = cell.style
        if (cell.link) win.links![addr] = cell.link
        if (cell.note) win.notes![addr] = cell.note
        if (cell.bar) win.bars![addr] = cell.bar
        if (cell.appearance) win.appearance![addr] = cell.appearance
      }
      win.text.push(text)
      win.values.push(values)
    }
    // Under a filter the rows are scattered: only one-row merges on shown rows stay, as the the app gives them.
    const area = { top: req.row, left: req.col, bottom: req.row + Math.max(rows, 1) - 1, right: req.col + Math.max(cols, 1) - 1 }
    const shown = new Set(sheetRows)
    win.merges = (store.sheet.merges ?? []).filter((m) => {
      const r = parseRange(m)
      if (!r) return false
      if (!list) return rectsMeet(r, area)
      return r.top === r.bottom && shown.has(r.top) && r.left <= area.right && r.right >= area.left
    })
    return win
  }

  /** Apply typed entries without an engine: `=` keeps the formula as a string, `1.5`-like text is a number, the rest text. */
  function apply(entries: CellEntry[]): CellUpdate[] {
    const changes: CellUpdate[] = []
    const touched = new Set<Store>()
    for (const e of entries) {
      const store = storeOf(e.sheet)
      const at = parseCell(e.cell)
      if (!store || !at) continue
      const addr = cellAddress(at.row, at.col)
      const cell: MemoryCell = { ...store.cells.get(addr) }
      delete cell.formula
      delete cell.text
      delete cell.color
      const t = e.text
      if (t === null || t === '') delete cell.value
      else if (t.startsWith('=') && t.length > 1) {
        cell.formula = t.slice(1)
        cell.value = null
      } else if (/^-?\d+(\.\d+)?$/.test(t)) cell.value = Number(t)
      else cell.value = t.startsWith("'") ? t.slice(1) : t
      if (Object.keys(cell).length) store.cells.set(addr, cell)
      else store.cells.delete(addr)
      touched.add(store)
      changes.push({ sheet: e.sheet, cell: addr, text: textOf(cell), value: cell.value ?? null, formula: cell.formula ?? null })
    }
    for (const store of touched) measure(store)
    return changes
  }

  const source: MemorySource = {
    maxCells: options.maxCells,
    sheets: (signal) => wait(() => stores.map(info), signal),
    window: (req, signal) => wait(() => windowOf(req), signal),
    edit(entries: CellEntry[]): Promise<EditOutcome> {
      return wait((): EditOutcome => {
        const failure = failures.shift()
        if (failure) return failure
        const unknown = entries.find((e) => !storeOf(e.sheet) || !parseCell(e.cell))
        if (unknown) return { ok: false, reason: 'failed', message: `No cell ${unknown.sheet}!${unknown.cell}` }
        return { ok: true, changes: apply(entries) }
      })
    },
    layout(change: LayoutChange): Promise<EditOutcome> {
      return wait((): EditOutcome => {
        const store = storeOf(change.sheet)
        if (!store) return { ok: false, reason: 'failed', message: `No sheet ${change.sheet}` }
        // Hidden is size 0; shown again a line takes the default, the book in memory keeps no size behind a hidden one.
        const put = (sizes: Map<number, number>, line: number, size: number | undefined) => {
          if (change.hidden) sizes.set(line, 0)
          else if (size !== undefined) sizes.set(line, size)
          else if (change.hidden === false && sizes.get(line) === 0) sizes.delete(line)
        }
        if ('columns' in change) for (const c of linesOf(change.columns, 'columns')) put(store.widths, c, change.width)
        else for (const r of linesOf(change.rows, 'rows')) put(store.heights, r, change.height)
        return { ok: true, changes: [] }
      })
    },
    find: (req, signal) => wait(() => findHits(query(storeOrThrow(req.sheet)), req), signal),
    replace: (req, by) =>
      wait((): ReplaceOutcome => {
        const store = storeOf(req.sheet)
        if (!store) return { ok: false, reason: 'failed', message: `No sheet ${req.sheet}`, replaced: 0 }
        const cells = replaceTexts(query(store), req, by)
        return { ok: true, changes: apply(cells.map((c) => ({ sheet: req.sheet, cell: cellAddress(c.row, c.col), text: c.text }))), replaced: cells.length }
      }),
    visibleRows: (sheet, filter, signal) =>
      wait(() => {
        const store = storeOrThrow(sheet)
        return { rows: visibleList(query(store), filter), total: store.maxRow }
      }, signal),
    columnValues: (sheet, col, filter, signal) => wait(() => columnList(query(storeOrThrow(sheet)), col, filter), signal),
    stats: (sheet, range, filter, signal) =>
      wait(() => {
        const rect = parseRange(range)
        if (!rect) throw new Error(`Not a range: ${range}`)
        return rangeStats(query(storeOrThrow(sheet)), rect, filter)
      }, signal),
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    failNext(...outcomes) {
      failures.push(...outcomes)
    },
    emit(event) {
      for (const l of listeners) l(event)
    },
    recompute(ms) {
      recomputing = { since: Date.now(), until: Date.now() + ms }
    },
    setCells(sheet, entries) {
      const changes = apply(Object.entries(entries).map(([cell, text]) => ({ sheet, cell, text })))
      source.emit({ kind: 'changed', changes })
      return changes
    },
  }
  return source
}
