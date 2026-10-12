// localSource: a sheet source over an xlcalc book whose session runs in a Web Worker pool or in the page's thread.
// The session holds the book, computes it and answers windows, sheets and edits as Excel makes them; this side keeps
// the sheets' shown values for search and the display filter, and the texts cells open with for Find in formulas and
// Replace, patched by the changes and dropped on a reload.
import { createInlineCalc, FUNCTION_NAMES, LOCALES, type AsyncSession, type BookInput, type Calc, type CellChange, type CellPatch, type EditResult, type LayoutEdit, type PatchFields, type RangeRef, type SheetExtent, type SheetLayout } from 'xlcalc'
import { cellAddress, linesKey, nameKey, parseCell, parseRange, rangeAddress, sameName } from 'xlcalc/a1'
import type { CellEntry, CellUpdate, EditOptions, EditOutcome, FindRequest, FormatPatch, LayoutChange, OutlineChange, ReplaceOutcome, SheetFilter, SheetInfo, SheetOp, SourceEvent, WindowRequest } from '../source'
import { loadLocale } from './locale'
import { outlineLayout, sheetOutline } from './outline'
import { boxesOf, cellKey, failed, needsReload, outcomeOf, refusalOf, updateOf, type CellsBox } from './outcome'
import { columnValuesOf, filterActive, filterKey, findIn, replaceIn, SheetValues, statsOf, visibleList } from './query'
import type { LocalSource, LocalSourceOptions } from './types'
import { alertRefusalOf, validationOf } from './validation'
import { windowOf } from './window'

const MAX_CELLS = 20_000
/** Cells of one window of the edit texts Find in formulas reads a sheet by. */
const EDIT_TEXT_CELLS = 20_000

const abortError = () => new DOMException('The request was aborted', 'AbortError')
function check(signal?: AbortSignal): void {
  if (signal?.aborted) throw abortError()
}

async function startCalc(opts: LocalSourceOptions): Promise<Calc> {
  const kind = typeof opts.calc === 'string' ? opts.calc : typeof Worker === 'function' ? 'pool' : 'inline'
  if (kind === 'inline') return createInlineCalc()
  const { createCalcPool } = await import('xlcalc/worker')
  return createCalcPool({ threads: 1, ...opts.pool })
}

/** A sheet of the book's layout with where its cells end, as the component lists it. */
function info(sheet: SheetLayout, used: SheetExtent): SheetInfo {
  const f = sheet.frozen
  const outline = sheetOutline(sheet)
  return {
    name: sheet.name,
    maxRow: used.rows,
    maxCol: used.columns,
    ...(f ? { freeze: { rows: f.rows ?? 0, cols: f.columns ?? 0 } } : {}),
    ...(sheet.hidden ? { hidden: sheet.hidden } : {}),
    ...(outline ? { outline } : {}),
  }
}

/** Lines as `setLayout` names them (`C`, `B:D`; `5`, `5:9`) from what the component sends (`C`, `C:C`, `3:4`). */
function linesOf(text: string, axis: 'columns' | 'rows'): string | null {
  const got = parseRange(text.includes(':') ? text : `${text}:${text}`)
  if (got?.kind !== axis) return null
  return axis === 'columns' ? linesKey(got.left, got.right, axis) : linesKey(got.top, got.bottom, axis)
}

/**
 * Cells scattered over a sheet (Replace All's) in rectangles of at most `max` cells, so that reading them back fits
 * windows and needs no reload; the cells come in Excel's order, which keeps neighbours together.
 */
function boxesUnder(sheet: string, cells: { row: number; col: number }[], max: number): CellsBox[] {
  const out: Required<CellsBox>[] = []
  let box: Required<CellsBox> | null = null
  for (const c of cells) {
    const addr = cellAddress(c.row, c.col)
    const r = box?.rect
    const rect = r ? { top: Math.min(r.top, c.row), left: Math.min(r.left, c.col), bottom: Math.max(r.bottom, c.row), right: Math.max(r.right, c.col) } : null
    if (box && rect && (rect.bottom - rect.top + 1) * (rect.right - rect.left + 1) <= max) {
      box.rect = rect
      box.only.add(addr)
    } else out.push((box = { sheet, rect: { top: c.row, left: c.col, bottom: c.row, right: c.col }, only: new Set([addr]) }))
  }
  return out
}

/** A source over `input`: the book goes to the session as it is (its cells are not copied), so do not change it afterwards. */
export function localSource(input: BookInput, opts: LocalSourceOptions = {}): LocalSource {
  const locale = input.locale ?? 'en-US'
  // Text sorts in the book's locale once it is loaded; an unknown locale fails the opening.
  let collator: Intl.Collator
  const maxCells = opts.maxCells ?? MAX_CELLS
  const listeners = new Set<(event: SourceEvent) => void>()
  /** Shown values of sheets that search or a filter needed, by the sheet's name key. */
  const values = new Map<string, Promise<SheetValues>>()
  /** The texts cells open with, of sheets Find in formulas or Replace needed, by the sheet's name key. */
  const editTexts = new Map<string, Promise<SheetValues>>()
  /** The list of rows of a sheet's display filter, as last given by `visibleRows`: windows under the filter read it. */
  const lists = new Map<string, { key: string; rows: number[] }>()
  let calc: Calc | null = null
  let session: AsyncSession | null = null

  /** The source's own changes one after another: each reads what the one before left. */
  let last: Promise<unknown> = Promise.resolve()
  function serial<A extends unknown[], R>(run: (...args: A) => Promise<R>): (...args: A) => Promise<R> {
    return (...args) => {
      const next = last.then(() => run(...args))
      last = next.catch(() => {})
      return next
    }
  }

  const ready = (async () => {
    await loadLocale(locale)
    calc = typeof opts.calc === 'object' ? opts.calc : await startCalc(opts)
    session = await calc.createSession(input, opts.session)
    collator = new Intl.Collator(locale, { numeric: true, sensitivity: 'base' })
  })()
  ready.catch(() => {})
  const open = async (): Promise<AsyncSession> => {
    await ready
    return session!
  }

  function emit(event: SourceEvent): void {
    for (const l of listeners) l(event)
  }

  /** The book changed in many places: kept values are dropped and subscribers read the sheets and windows again. */
  function reload(): void {
    values.clear()
    editTexts.clear()
    lists.clear()
    emit({ kind: 'reload' })
  }
  /** Whether a conditional format of the book shows numbers by a format of its own, which changes cells' shown text. */
  const rulesFormat = input.sheets.some((sh) => sh.conditionalFormats?.some((r) => 'style' in r && r.style?.numberFormat !== undefined))
  /**
   * Conditional formats may show cells otherwise anywhere their rules meet (the library's `repaint`), which cell updates do
   * not carry: subscribers read the windows over those ranges again. A rule changes no value, so the kept values stay unless
   * a rule formats numbers (reading a sheet's values again is a pass over all its cells). The display filters' lists stay:
   * the component keeps its list until the filter is applied again, as Excel keeps rows until Reapply.
   */
  function repaint(ranges: RangeRef[]): void {
    if (rulesFormat) values.clear()
    emit({ kind: 'repaint', ranges: ranges.map((x) => ({ sheet: x.sheet, range: x.range })) })
  }

  /** The session's changes as cell updates, the edited cells' in their place, and the kept values brought up to date. */
  function take(changes: CellChange[], edited: Map<string, CellUpdate> = new Map()): CellUpdate[] {
    const out = new Map(edited)
    for (const ch of changes) {
      const key = cellKey(ch.sheet, ch.cell)
      if (!out.has(key)) out.set(key, { sheet: ch.sheet, cell: ch.cell, text: ch.text ?? '', value: ch.value, ...(ch.color ? { color: ch.color } : {}) })
    }
    const list = [...out.values()]
    for (const u of list) {
      const at = parseCell(u.cell)
      if (!at) continue
      void values.get(nameKey(u.sheet))?.then((v) => v.set(at.row, at.col, u.text, u.value))
      // only an edited cell opens otherwise: a recomputed one keeps its formula
      if (u.editText !== undefined) void editTexts.get(nameKey(u.sheet))?.then((t) => t.set(at.row, at.col, u.editText!, undefined))
    }
    return list
  }

  /** A sheet's shown values, read once over where its cells end. */
  function sheetValues(s: AsyncSession, sheet: string): Promise<SheetValues> {
    const key = nameKey(sheet)
    let got = values.get(key)
    if (!got) {
      got = s.getUsedRange(sheet).then(async (used) => new SheetValues(used.rows && used.columns ? await s.getCells(sheet, `A1:${cellAddress(used.rows, used.columns)}`) : {}))
      values.set(key, got)
      got.catch(() => values.delete(key))
    }
    return got
  }

  /**
   * The text each cell of a sheet opens with (a formula in the book's locale, a value by the format it shows), read once
   * over where its cells end in windows of `EDIT_TEXT_CELLS`, so that other calls of the session go between them.
   */
  function sheetEditTexts(s: AsyncSession, sheet: string): Promise<SheetValues> {
    const key = nameKey(sheet)
    let got = editTexts.get(key)
    if (!got) {
      got = (async () => {
        const used = await s.getUsedRange(sheet)
        const texts = new SheetValues()
        const step = Math.max(1, Math.floor(EDIT_TEXT_CELLS / Math.max(1, used.columns)))
        for (let top = 1; top <= used.rows && used.columns; top += step) {
          const range = `${cellAddress(top, 1)}:${cellAddress(Math.min(used.rows, top + step - 1), used.columns)}`
          const v = await s.getViewport(sheet, range, { editText: true, maxCells: step * used.columns })
          for (const [addr, cell] of Object.entries(v.cells)) {
            const at = cell.editText ? parseCell(addr) : null
            if (at) texts.set(at.row, at.col, cell.editText!, undefined)
          }
        }
        return texts
      })()
      editTexts.set(key, got)
      got.catch(() => editTexts.delete(key))
    }
    return got
  }

  /** The rows of a sheet under a filter: the list `visibleRows` gave for it, so windows match what the component holds. */
  async function rowsUnder(s: AsyncSession, sheet: string, filter: SheetFilter, fresh: boolean): Promise<number[]> {
    const key = filterKey(filter)
    const kept = lists.get(nameKey(sheet))
    if (!fresh && kept?.key === key) return kept.rows
    const [v, used] = await Promise.all([sheetValues(s, sheet), s.getUsedRange(sheet)])
    const rows = visibleList(v, filter, used.rows, collator)
    lists.set(nameKey(sheet), { key, rows })
    return rows
  }

  /** The edited cells as they are now, with their formulas as shown and edit texts; `truncated` when a sheet's rectangle held more than a window. */
  async function editedIn(s: AsyncSession, boxes: CellsBox[]): Promise<{ updates: Map<string, CellUpdate>; truncated: boolean }> {
    const updates = new Map<string, CellUpdate>()
    let truncated = false
    await Promise.all(
      boxes.map(async (b) => {
        const v = await s.getViewport(b.sheet, rangeAddress(b.rect), { formulaLocal: true, editText: true, maxCells })
        truncated ||= v.truncated
        const got = parseRange(v.range)!
        for (let r = got.top; r <= got.bottom; r++) {
          for (let c = got.left; c <= got.right; c++) {
            const addr = cellAddress(r, c)
            if (!b.only || b.only.has(addr)) updates.set(cellKey(b.sheet, addr), updateOf(b.sheet, addr, v.cells[addr]))
          }
        }
      }),
    )
    return { updates, truncated }
  }

  /** A change made: its updates, the edited cells' as they are now, told to subscribers, then the looks it may have changed; a reload where updates cannot carry it. */
  async function settle(s: AsyncSession, got: EditResult, boxes: CellsBox[]): Promise<EditOutcome> {
    const { updates, truncated } = await editedIn(s, boxes)
    const changes = take(got.changes, updates)
    if (truncated || needsReload(got)) reload()
    else {
      if (changes.length) emit({ kind: 'changed', changes })
      if (got.repaint?.length) repaint(got.repaint)
    }
    return { ok: true, changes }
  }

  /**
   * Typed texts, each read as typed into its cell, in one step; only empty cells (Delete over a selection) clear the content.
   * A value a rule of data validation does not take, of any style, refuses the call with the rule's alert before anything is
   * written, as Excel's alert holds the entry; the library picks the strictest rule a batch breaks. With `accept` (a
   * warning's Yes, an information's OK) the texts are written without asking the rules again.
   */
  async function edit(entries: CellEntry[], opts?: EditOptions): Promise<EditOutcome> {
    const s = await open()
    const cells: { sheet: string; cell: string; text: string | null }[] = []
    for (const e of entries) {
      const at = parseCell(e.cell)
      if (!at) return failed(`No cell ${e.sheet}!${e.cell}`)
      cells.push({ sheet: e.sheet, cell: cellAddress(at.row, at.col), text: e.text })
    }
    if (!cells.length) return { ok: true, changes: [] }
    const boxes = boxesOf(cells)
    try {
      if (cells.every((c) => c.text === null)) return await settle(s, await s.clearCells(cells.map((c) => ({ sheet: c.sheet, range: c.cell })), 'contents'), boxes)
      const texts = cells.map((c) => ({ sheet: c.sheet, cell: c.cell, text: c.text ?? '' }))
      const got = await s.setCellTexts(texts, { refuseInvalid: !opts?.accept, ...(opts?.inPlace ? { inPlace: true } : {}) })
      if (got.ok) return await settle(s, got, boxes)
      return got.code === 'validation' ? alertRefusalOf(got) : refusalOf(got)
    } catch (e) {
      return outcomeOf(e, 'refused')
    }
  }

  /** Clipboard text as Excel pastes it at a cell, one step of the undo; the selection goes over what it filled. */
  async function paste(sheet: string, cell: string, text: string): Promise<EditOutcome> {
    const s = await open()
    try {
      const got = await s.pasteText(sheet, cell, text)
      if (!got.ok) return refusalOf(got)
      const out = await settle(s, got, [{ sheet, rect: parseRange(got.range)! }])
      return out.ok ? { ...out, select: { sheet, range: got.range } } : out
    } catch (e) {
      return outcomeOf(e, 'refused')
    }
  }

  /** Control+Return: the text read once in the active cell and entered into every cell of the range, which stays selected; the active cell's rule refuses as in `edit`. */
  async function fill(sheet: string, range: string, text: string, active: string, opts?: EditOptions): Promise<EditOutcome> {
    const s = await open()
    const rect = parseRange(range)
    if (!rect) return failed(`Not a range: ${range}`)
    try {
      const got = await s.fillCellText(sheet, range, text, { active, refuseInvalid: !opts?.accept, ...(opts?.inPlace ? { inPlace: true } : {}) })
      if (!got.ok) return got.code === 'validation' ? alertRefusalOf(got) : refusalOf(got)
      const out = await settle(s, got, [{ sheet, rect }])
      return out.ok ? { ...out, select: { sheet, range: rangeAddress(rect) } } : out
    } catch (e) {
      return outcomeOf(e, 'refused')
    }
  }

  /** The session's undo or redo, whole; the selection goes over the cells it took back, on the first of their sheets. */
  async function step(kind: 'undo' | 'redo'): Promise<EditOutcome> {
    const s = await open()
    try {
      const got = await s[kind]()
      if (!got) return { ok: true, changes: [] }
      // The display filters' lists stay as the component holds them; a step that needs more reloads, which drops them.
      values.clear()
      // TODO(xlcalc): a structure step's undo does not list the cells it empties (a copy onto empty cells), so the sheet is read again, as after the operation.
      if (got.call === 'editStructure') {
        reload()
        return { ok: true, changes: take(got.changes) }
      }
      const boxes = boxesOf(got.cells ?? [])
      const out = await settle(s, got, boxes)
      const first = boxes[0]
      return out.ok && first ? { ...out, select: { sheet: first.sheet, range: rangeAddress(first.rect) } } : out
    } catch (e) {
      return outcomeOf(e, 'failed')
    }
  }

  async function structure(op: SheetOp): Promise<EditOutcome> {
    const s = await open()
    try {
      const got = await s.editStructure(op)
      if (!got.ok) return { ok: false, reason: 'refused', message: got.message }
      const changes = take(got.changes)
      reload()
      return { ok: true, changes }
    } catch (e) {
      return outcomeOf(e, 'failed')
    }
  }

  async function setLayout(change: LayoutChange): Promise<EditOutcome> {
    const s = await open()
    try {
      const shown = (size: number | undefined) => (change.hidden !== undefined ? { hidden: change.hidden ? true : null } : size !== undefined && size > 0 ? { hidden: null } : {})
      let edit: LayoutEdit
      if ('columns' in change) {
        const columns = linesOf(change.columns, 'columns')
        if (!columns) return failed(`Not columns: ${change.columns}`)
        edit = change.fit ? { sheet: change.sheet, columns, fit: true } : { sheet: change.sheet, columns, set: { ...(change.width !== undefined ? { width: change.width } : {}), ...shown(change.width) } }
      } else {
        const rows = linesOf(change.rows, 'rows')
        if (!rows) return failed(`Not rows: ${change.rows}`)
        edit = { sheet: change.sheet, rows, set: { ...(change.height !== undefined ? { height: change.height } : {}), ...shown(change.height) } }
      }
      const got = await s.setLayout([edit])
      const changes = take(got.changes)
      // A size set with the mouse the component already shows; lines hidden, shown or fitted it reads again.
      if (change.hidden !== undefined || ('fit' in change && change.fit)) reload()
      else if (changes.length) emit({ kind: 'changed', changes })
      return { ok: true, changes }
    } catch (e) {
      return outcomeOf(e, 'failed')
    }
  }

  /** Formatting as Excel formats a selection: the style over the one each cell shows, in one patch of the session. */
  async function format(sheet: string, patches: FormatPatch[]): Promise<EditOutcome> {
    const s = await open()
    const cells: CellPatch[] = []
    for (const p of patches) {
      const set: PatchFields = {}
      if (p.style !== undefined) set.style = p.style
      if (p.numberFormat !== undefined) set.numberFormat = p.numberFormat
      if (Object.keys(set).length) cells.push({ sheet, range: p.range, set })
    }
    if (!cells.length) return { ok: true, changes: [] }
    try {
      const got = await s.patchCells(cells)
      const changes = take(got.changes)
      // A look, a number format and an edit text are the window's to carry, not a cell update's.
      reload()
      return { ok: true, changes }
    } catch (e) {
      return outcomeOf(e, 'refused')
    }
  }

  /**
   * Excel's Replace All: the request's hits in the texts cells open with, whatever `lookIn` says, each written as typed with
   * its matches replaced, in one `setCellTexts` and so one step of undo. As Excel's paste does, a replacement neither widens
   * a column nor checks data validation (not measured in Excel). A text the session refuses refuses the call, nothing written.
   */
  async function replace(req: FindRequest, by: string): Promise<ReplaceOutcome> {
    const s = await open()
    try {
      const [texts, rows] = await Promise.all([sheetEditTexts(s, req.sheet), filterActive(req.filter) ? rowsUnder(s, req.sheet, req.filter, false) : null])
      const cells = replaceIn(texts, req, by, rows ? new Set(rows) : null)
      if (!cells.length) return { ok: true, changes: [], replaced: 0 }
      const got = await s.setCellTexts(
        cells.map((c) => ({ sheet: req.sheet, cell: cellAddress(c.row, c.col), text: c.text })),
        { widen: false },
      )
      if (!got.ok) return { ...refusalOf(got), replaced: 0 }
      return { ...(await settle(s, got, boxesUnder(req.sheet, cells, maxCells))), replaced: cells.length }
    } catch (e) {
      return { ...outcomeOf(e, 'refused'), replaced: 0 }
    }
  }

  /** Excel's outline commands (`./outline`): one `setLayout` call, one step of the undo, then the sheets and windows are read again. */
  async function outline(change: OutlineChange): Promise<EditOutcome> {
    const s = await open()
    try {
      const sheet = (await s.getLayout()).sheets.find((x) => sameName(x.name, change.sheet))
      if (!sheet) return failed(`No sheet ${change.sheet}`)
      const got = outlineLayout(sheet, change)
      if (!got.ok) return { ok: false, reason: 'refused', message: got.message }
      if (!got.edits.length) return { ok: true, changes: [] }
      const changes = take((await s.setLayout(got.edits)).changes)
      reload()
      return { ok: true, changes }
    } catch (e) {
      return outcomeOf(e, 'failed')
    }
  }

  async function freeze(sheet: string, panes: { rows: number; cols: number } | null): Promise<EditOutcome> {
    const s = await open()
    const frozen = panes && (panes.rows > 0 || panes.cols > 0) ? { ...(panes.rows > 0 ? { rows: panes.rows } : {}), ...(panes.cols > 0 ? { columns: panes.cols } : {}) } : null
    try {
      const got = await s.setLayout([{ sheet, frozen }])
      const changes = take(got.changes)
      reload()
      return { ok: true, changes }
    } catch (e) {
      return outcomeOf(e, 'failed')
    }
  }

  const source: LocalSource = {
    maxCells,
    locale,
    ready,
    async sheets(signal) {
      const s = await open()
      const { sheets } = await s.getLayout()
      const used = await Promise.all(sheets.map((x) => s.getUsedRange(x.name)))
      check(signal)
      return sheets.map((x, i) => info(x, used[i]!))
    },
    async window(req: WindowRequest, signal) {
      const s = await open()
      check(signal)
      const cols = Math.max(0, req.cols)
      const count = Math.max(0, Math.min(req.rows, Math.floor(maxCells / Math.max(1, cols))))
      const filtered = filterActive(req.filter)
      let rows: number[]
      if (filtered) rows = (await rowsUnder(s, req.sheet, req.filter!, false)).slice(req.row - 1, req.row - 1 + count)
      else rows = Array.from({ length: count }, (_, i) => req.row + i)
      // the viewport takes rows in ascending order; the window keeps the display order
      const asked = [...rows].sort((a, b) => a - b)
      const top = asked[0] ?? 1
      const range = `${cellAddress(top, req.col)}:${cellAddress(asked[asked.length - 1] ?? top, req.col + Math.max(1, cols) - 1)}`
      const v = await s.getViewport(req.sheet, range, { formulaLocal: true, editText: true, maxCells, ...(filtered || !rows.length ? { rows: asked } : {}) })
      check(signal)
      return windowOf(v, { row: req.row, rows, col: req.col, cols, filtered, notice: opts.notice })
    },
    edit: serial(edit),
    paste: serial(paste),
    fill: serial(fill),
    undo: serial(() => step('undo')),
    redo: serial(() => step('redo')),
    structure: serial(structure),
    layout: serial(setLayout),
    format: serial(format),
    replace: serial(replace),
    outline: serial(outline),
    freeze: serial(freeze),
    get history() {
      return session?.history ?? { undo: 0, redo: 0 }
    },
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    async validation(sheet, cell, signal) {
      const s = await open()
      check(signal)
      const got = validationOf(await s.getValidation(sheet, cell))
      check(signal)
      return got
    },
    /** Every function a formula can call, by its name in the book's locale (English where the locale names none, as Excel ru-RU shows `ABS`), in the locale's alphabetical order. */
    async functionNames(signal) {
      await ready
      check(signal)
      const local = LOCALES.get(locale)?.functions ?? {}
      return FUNCTION_NAMES.map((name) => local[name] ?? name).sort(new Intl.Collator(locale).compare)
    },
    async explain(sheet, cell) {
      return (await (await open()).explain(sheet, cell)).text
    },
    async stats(sheet, range, filter, signal) {
      const s = await open()
      const [views, rows] = await Promise.all([s.getCells(sheet, range), filterActive(filter) ? rowsUnder(s, sheet, filter, false) : null])
      check(signal)
      return statsOf(views, rows ? new Set(rows) : null)
    },
    async find(req, signal) {
      const s = await open()
      const texts = req.lookIn === 'values' ? sheetValues(s, req.sheet) : sheetEditTexts(s, req.sheet)
      const [v, rows] = await Promise.all([texts, filterActive(req.filter) ? rowsUnder(s, req.sheet, req.filter, false) : null])
      check(signal)
      return findIn(v, req, rows ? new Set(rows) : null)
    },
    async visibleRows(sheet, filter, signal) {
      const s = await open()
      if (!filterActive(filter)) {
        lists.delete(nameKey(sheet))
        return { rows: null, total: (await s.getUsedRange(sheet)).rows }
      }
      const [rows, used] = await Promise.all([rowsUnder(s, sheet, filter, true), s.getUsedRange(sheet)])
      check(signal)
      return { rows, total: used.rows }
    },
    async columnValues(sheet, col, filter, signal) {
      const s = await open()
      const [v, used] = await Promise.all([sheetValues(s, sheet), s.getUsedRange(sheet)])
      check(signal)
      return columnValuesOf(v, col, filter, used.rows, collator)
    },
    async book() {
      return (await open()).getBook()
    },
    async snapshot() {
      return (await open()).getSnapshot()
    },
    async close() {
      await ready.catch(() => {})
      await session?.close()
      if (calc && typeof opts.calc !== 'object') await calc.close()
    },
  }
  return source
}
