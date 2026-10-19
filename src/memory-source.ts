// A source over an in-memory workbook, for tests and examples: values, displayed text, formulas as strings and every display field; no engine.
// It has every optional method of 0.2.0 (search, filter, column values, totals), Replace All, frozen panes, a precedent
// tree from the formulas as given, and can simulate a workbook being recalculated.
import { cellAddress, outlineGroups, parseCell, parseRange, quoteSheetIfNeeded, rangeAddress, rectsMeet, sameName } from 'xlcalc/a1'
import type { CellStyle, CellValue, ConditionalAppearance } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'
import { loadEntry, type Entry } from './entry'
import type { CellEntry, CellUpdate, EditOptions, EditOutcome, LayoutChange, OutlineChange, ReplaceOutcome, SheetFilter, SheetInfo, SheetPending, SheetSource, SheetWindow, SourceEvent, WindowRequest } from './contract/source'
import { columnList, findHits, rangeStats, replaceTexts, visibleList, type MemorySheetQuery } from './memory-query'
import { axisOf, lineSpan, type LineAxis } from './lines'
import { outlineCommand, outlineOf, outlineSheet, type AxisLines } from './memory-outline'
import { abortError } from './abort'
import { filterKey } from './filter-key'
import { failedOutcome } from './failure'
import { isGridRef } from './grid-ref'
import { DEFAULT_MAX_CELLS, rowsFor } from './window-size'
import { MemoryRules, memoryRefusal, PLAIN_NUMBER } from './memory-validation'
import type { MemoryBook, MemoryCell, MemorySheet, MemorySource, MemorySourceOptions, MemoryValidation } from './contract/memory'

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

/** A cell's value: its own, otherwise the text it was given alone, otherwise empty. */
const valueOf = (cell: MemoryCell | undefined): CellValue => cell?.value ?? (cell?.text || null)

interface Store {
  sheet: MemorySheet
  cells: Map<string, MemoryCell>
  widths: Map<number, number>
  heights: Map<number, number>
  rules: MemoryRules
  /** Outline levels and collapsed flags by row or column. */
  levels: Record<LineAxis, Map<number, number>>
  marks: Record<LineAxis, Set<number>>
  /** The sheet's frozen panes, as `freeze` leaves them. */
  freeze?: { rows: number; cols: number }
  maxRow: number
  maxCol: number
}

/** The levels of precedents and the nodes `explain` writes, as xlcalc's defaults. */
const EXPLAIN_DEPTH = 3
const EXPLAIN_NODES = 200

/** Each column or row of a key as a workbook accepts it (`B`, `B:D`, `C:C`; `5`, `5:9`); none for a key it does not accept. */
function linesOf(key: string, axis: LineAxis): number[] {
  const span = lineSpan(key, axis)
  const out: number[] = []
  if (span) for (let i = span.first; i <= span.last; i++) out.push(i)
  return out
}

/** A source over `book`; the workbook's objects are copied, and edits change the copy. */
export function memorySource(book: MemoryBook, options: MemorySourceOptions = {}): MemorySource {
  const delay = options.delay ?? 0
  const maxCells = options.maxCells ?? DEFAULT_MAX_CELLS
  const styles = book.styles ?? []
  const functions = book.functions?.slice()
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
    const levels = { rows: new Map<number, number>(), columns: new Map<number, number>() }
    const marks = { rows: new Set<number>(), columns: new Set<number>() }
    for (const axis of ['rows', 'columns'] as const) {
      for (const [key, level] of Object.entries(sheet.outline?.[axis] ?? {})) for (const i of linesOf(key, axis)) levels[axis].set(i, level)
      for (const key of sheet.outline?.collapsed ?? []) for (const i of linesOf(key, axis)) marks[axis].add(i)
    }
    const store: Store = { sheet, cells, widths, heights, rules: new MemoryRules(sheet.validations), levels, marks, freeze: sheet.freeze, maxRow: 0, maxCol: 0 }
    measure(store)
    return store
  })
  const listeners = new Set<(event: SourceEvent) => void>()
  /** Each sheet's rows under its display filter as `visibleRows` last returned them: windows under that filter use them, so a row that no longer matches after an edit stays until the filter is reapplied, as the component holds it and `./local` reads it. */
  const lists = new Map<Store, { key: string; rows: number[] }>()
  const failures: Extract<EditOutcome, { ok: false }>[] = []
  /** A simulated recalculation: when it started and when it ends. */
  let recomputing: { since: number; until: number } | null = null

  /** A sheet as the queries read it: displayed text and values by row and column. */
  function query(store: Store): MemorySheetQuery {
    return {
      maxRow: store.maxRow,
      maxCol: store.maxCol,
      textAt: (r, c) => {
        const cell = store.cells.get(cellAddress(r, c))
        return cell ? textOf(cell) : ''
      },
      valueAt: (r, c) => valueOf(store.cells.get(cellAddress(r, c))),
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

  /** The sheet's extent: non-empty cells, and merges. */
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
  /** xlcalc's typed-entry module once a search has loaded it: Find and Replace search what cells open with. */
  let entry: Entry | null = null
  /** `wait` once xlcalc/entry has loaded. */
  const withEntry = <T>(make: () => T, signal?: AbortSignal): Promise<T> =>
    loadEntry().then((m) => {
      entry = m
      return wait(make, signal)
    })
  /** What a cell opens with, as `apply` reads it back: `=` and its formula, otherwise the value's `entryText` in en-US (text with Excel's apostrophe where it would be parsed as a value). */
  function editTextOf(cell: MemoryCell): string {
    if (cell.formula !== undefined) return `=${cell.formula}`
    if (cell.value === undefined || cell.value === null) return cell.text ?? ''
    // withEntry has loaded it for every caller; plain text is only the type's fallback
    return entry ? entry.entryText(cell.value, cell.numberFormat ? { numberFormat: cell.numberFormat } : undefined) : plainText(cell.value)
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
    const { name, hidden, locked, columns, rowNumbers, totals } = store.sheet
    const outline = outlineOf(outlineGroups(outlineSheet({ rows: axisLines(store, 'rows'), columns: axisLines(store, 'columns') }, store.sheet.outline)), store.sheet.outline)
    return { name, maxRow: store.maxRow, maxCol: store.maxCol, freeze: store.freeze, hidden, locked, columns, rowNumbers, totals, ...(outline ? { outline } : {}) }
  }

  /**
   * A cell's precedents as xlcalc's `explain` writes them in en-US, from the formulas as given (there is no engine): a node
   * per line with its value and formula, each two-space indent one level, three levels, and a range as one node with its non-empty cells.
   */
  function explainOf(sheet: string, cell: string): string {
    const lines: string[] = []
    const value = (v: CellValue) => (v === null ? 'empty' : typeof v === 'object' ? plainText(v) : JSON.stringify(v))
    const visit = (store: Store, addr: string, depth: number, indent: string) => {
      const c = store.cells.get(addr)
      const name = quoteSheetIfNeeded(store.sheet.name)
      lines.push(`${indent}${name}!${addr} = ${value(valueOf(c))}${c?.formula !== undefined ? `  ← ${c.formula}` : ''}`)
      if (c?.formula === undefined || depth === 0) return
      for (const r of scanReferences(c.formula).filter(isGridRef)) {
        if (lines.length >= EXPLAIN_NODES) return
        const target = r.sheet === undefined ? store : stores.find((x) => sameName(x.sheet.name, r.sheet!))
        if (!target) continue
        if (r.top === r.bottom && r.left === r.right) visit(target, cellAddress(r.top, r.left), depth - 1, `${indent}  `)
        else {
          let filled = 0
          for (const key of target.cells.keys()) {
            const at = parseCell(key)
            if (at && at.row >= r.top && at.row <= r.bottom && at.col >= r.left && at.col <= r.right) filled++
          }
          lines.push(`${indent}  ${quoteSheetIfNeeded(target.sheet.name)}!${rangeAddress(r)} (${filled} cells)`)
        }
      }
    }
    const at = parseCell(cell)
    if (at) visit(storeOrThrow(sheet), cellAddress(at.row, at.col), EXPLAIN_DEPTH, '')
    return lines.join('\n')
  }

  /** A sheet's rows or columns on one axis as the outline reads them: their levels and flags, and those hidden by a size of 0. */
  function axisLines(store: Store, axis: LineAxis): AxisLines {
    const hidden = new Set<number>()
    for (const [i, size] of axis === 'rows' ? store.heights : store.widths) if (size === 0) hidden.add(i)
    return { levels: store.levels[axis], marks: store.marks[axis], hidden }
  }

  /** The rows under a filter: the list `visibleRows` returned for it, otherwise the sheet's rows as they are; null without conditions. */
  function listUnder(store: Store, filter: SheetFilter | undefined): number[] | null {
    const kept = filter ? lists.get(store) : undefined
    return kept && kept.key === filterKey(filter!) ? kept.rows : visibleList(query(store), filter)
  }

  /** A range; under a filter `row` is a display index and the rows come from the filter's list. */
  function windowOf(req: WindowRequest): SheetWindow | SheetPending {
    if (recomputing && Date.now() < recomputing.until) return { pending: { elapsedMs: Date.now() - recomputing.since } }
    const store = storeOrThrow(req.sheet)
    const list = listUnder(store, req.filter)
    const cols = Math.max(0, req.cols)
    const limit = list ? Math.max(0, list.length - req.row + 1) : Infinity
    const rows = Math.min(rowsFor(maxCells, req.rows, cols), limit)
    const win: SheetWindow = {
      row: req.row,
      col: req.col,
      text: [],
      values: [],
      formulas: {},
      colors: {},
      styleIds: {},
      styles,
      numberFormats: {},
      widths: [],
      rowHeights: {},
      merges: [],
      defaultRowHeight: store.sheet.defaultRowHeight,
      tables: store.sheet.tables ?? [],
      appearance: {},
      links: {},
      notes: {},
      drill: {},
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
        values.push(valueOf(cell))
        if (!cell) continue
        if (cell.formula !== undefined) win.formulas[addr] = cell.formula
        if (cell.color) win.colors[addr] = cell.color
        if (cell.style !== undefined) win.styleIds[addr] = cell.style
        if (cell.numberFormat) win.numberFormats![addr] = cell.numberFormat
        if (cell.link) win.links![addr] = cell.link
        if (cell.note) win.notes![addr] = cell.note
        if (cell.drill) win.drill![addr] = true
        if (cell.appearance) win.appearance![addr] = cell.appearance
      }
      win.text.push(text)
      win.values.push(values)
    }
    // Under a filter the rows are scattered: only one-row merges on visible rows stay, as a display filter returns them.
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

  /** Applies typed entries without an engine: `=` keeps the formula as a string, `1.5`-like text is a number, the rest is text. */
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
      } else if (PLAIN_NUMBER.test(t)) cell.value = Number(t)
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
    maxCells,
    sheets: (signal) => wait(() => stores.map(info), signal),
    window: (req, signal) => wait(() => windowOf(req), signal),
    /** Typed entries; a value that violates a data validation rule rejects them all with its alert, unless `accept` keeps it. */
    edit(entries: CellEntry[], opts?: EditOptions): Promise<EditOutcome> {
      return wait((): EditOutcome => {
        const failure = failures.shift()
        if (failure) return failure
        const unknown = entries.find((e) => !storeOf(e.sheet) || !parseCell(e.cell))
        if (unknown) return { ok: false, reason: 'failed', message: `No cell ${unknown.sheet}!${unknown.cell}` }
        const refused = opts?.accept ? null : memoryRefusal(entries, (sheet) => storeOf(sheet)?.rules)
        return refused ?? { ok: true, changes: apply(entries) }
      })
    },
    validation: (sheet, cell, signal) => wait(() => storeOrThrow(sheet).rules.show(cell), signal),
    ...(functions ? { functionNames: (signal?: AbortSignal) => wait(() => [...functions], signal) } : {}),
    /** Frozen panes as the in-memory workbook keeps them; nothing frozen unfreezes. */
    freeze: (sheet, panes) =>
      wait((): EditOutcome => {
        const store = storeOf(sheet)
        if (!store) return { ok: false, reason: 'failed', message: `No sheet ${sheet}` }
        store.freeze = panes && (panes.rows > 0 || panes.cols > 0) ? { rows: Math.max(0, panes.rows), cols: Math.max(0, panes.cols) } : undefined
        return { ok: true, changes: [] }
      }),
    explain: (sheet, cell) => wait(() => explainOf(sheet, cell)),
    layout(change: LayoutChange): Promise<EditOutcome> {
      return wait((): EditOutcome => {
        const store = storeOf(change.sheet)
        if (!store) return { ok: false, reason: 'failed', message: `No sheet ${change.sheet}` }
        // Hidden is a size of 0; when unhidden, a row or column gets the default size, as the in-memory workbook keeps no size behind a hidden one.
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
    /** Excel's outline commands as xlcalc's `outlineEdit` runs them; a hidden row or column has a size of 0, and when unhidden it gets the default. */
    outline(change: OutlineChange): Promise<EditOutcome> {
      return wait((): EditOutcome => {
        const store = storeOf(change.sheet)
        if (!store) return { ok: false, reason: 'failed', message: `No sheet ${change.sheet}` }
        const axis = axisOf(change)
        const lines = axisLines(store, axis)
        let refused: string | null
        try {
          refused = outlineCommand(lines, store.sheet.outline, change)
        } catch (e) {
          return failedOutcome(e)
        }
        if (refused !== null) return { ok: false, reason: 'refused', message: refused }
        const sizes = axis === 'rows' ? store.heights : store.widths
        for (const i of lines.hidden) sizes.set(i, 0)
        for (const [i, size] of sizes) if (size === 0 && !lines.hidden.has(i)) sizes.delete(i)
        return { ok: true, changes: [] }
      })
    },
    find: (req, signal) => withEntry(() => findHits(query(storeOrThrow(req.sheet)), req), signal),
    replace: (req, by) =>
      withEntry((): ReplaceOutcome => {
        const store = storeOf(req.sheet)
        if (!store) return { ok: false, reason: 'failed', message: `No sheet ${req.sheet}`, replaced: 0 }
        const cells = replaceTexts(query(store), req, by)
        return { ok: true, changes: apply(cells.map((c) => ({ sheet: req.sheet, cell: cellAddress(c.row, c.col), text: c.text }))), replaced: cells.length }
      }),
    visibleRows: (sheet, filter, signal) =>
      wait(() => {
        const store = storeOrThrow(sheet)
        const rows = visibleList(query(store), filter)
        if (rows) lists.set(store, { key: filterKey(filter), rows })
        else lists.delete(store)
        return { rows }
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
