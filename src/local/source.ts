// localSource: a sheet source over an xlcalc book whose session runs in a Web Worker pool or in the page's thread.
// The session holds the book, computes it and answers windows, sheets and edits as Excel makes them; this side keeps
// the sheets' shown values for search and the display filter, patched by the changes and dropped on a reload.
import { createInlineCalc, type AsyncSession, type BookInput, type Calc, type CellChange, type CellPatch, type EditResult, type LayoutEdit, type PatchFields, type RangeRef, type SheetExtent, type SheetLayout } from 'xlcalc'
import { cellAddress, linesKey, nameKey, parseCell, parseRange, rangeAddress, sameName } from 'xlcalc/a1'
import type { CellEntry, CellUpdate, EditOutcome, FormatPatch, LayoutChange, OutlineChange, SheetFilter, SheetInfo, SheetOp, SourceEvent, WindowRequest } from '../source'
import { loadLocale } from './locale'
import { outlineLayout, sheetOutline } from './outline'
import { boxesOf, cellKey, failed, needsReload, outcomeOf, refusalOf, updateOf, type CellsBox } from './outcome'
import { columnValuesOf, filterActive, filterKey, findIn, SheetValues, statsOf, visibleList } from './query'
import type { LocalSource, LocalSourceOptions } from './types'
import { windowOf } from './window'

const MAX_CELLS = 20_000

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

/** A source over `input`: the book goes to the session as it is (its cells are not copied), so do not change it afterwards. */
export function localSource(input: BookInput, opts: LocalSourceOptions = {}): LocalSource {
  const locale = input.locale ?? 'en-US'
  // Text sorts in the book's locale once it is loaded; an unknown locale fails the opening.
  let collator: Intl.Collator
  const maxCells = opts.maxCells ?? MAX_CELLS
  const listeners = new Set<(event: SourceEvent) => void>()
  /** Shown values of sheets that search or a filter needed, by the sheet's name key. */
  const values = new Map<string, Promise<SheetValues>>()
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
      if (at) void values.get(nameKey(u.sheet))?.then((v) => v.set(at.row, at.col, u.text, u.value))
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
   * A value a stop rule of data validation does not take refuses the call before anything is written, as Excel's alert
   * holds the entry; a warning's or an information's is written, as their Yes and OK keep it (the sheet has no alert yet).
   * The library refuses a batch by its strictest rule, so a warning refusal means no cell breaks a stop rule.
   */
  async function edit(entries: CellEntry[]): Promise<EditOutcome> {
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
      let got = await s.setCellTexts(texts, { refuseInvalid: true })
      if (!got.ok && got.code === 'validation' && got.severity !== 'stop') got = await s.setCellTexts(texts)
      return got.ok ? await settle(s, got, boxes) : refusalOf(got)
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

  /** Control+Return: the text read once in the active cell and entered into every cell of the range, which stays selected; its rule refuses as in `edit`. */
  async function fill(sheet: string, range: string, text: string, active: string): Promise<EditOutcome> {
    const s = await open()
    const rect = parseRange(range)
    if (!rect) return failed(`Not a range: ${range}`)
    try {
      let got = await s.fillCellText(sheet, range, text, { active, refuseInvalid: true })
      if (!got.ok && got.code === 'validation' && got.severity !== 'stop') got = await s.fillCellText(sheet, range, text, { active })
      if (!got.ok) return refusalOf(got)
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
    outline: serial(outline),
    freeze: serial(freeze),
    get history() {
      return session?.history ?? { undo: 0, redo: 0 }
    },
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
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
    async find({ sheet, query, filter }, signal) {
      const s = await open()
      const [v, rows] = await Promise.all([sheetValues(s, sheet), filterActive(filter) ? rowsUnder(s, sheet, filter, false) : null])
      check(signal)
      return findIn(v, query, rows ? new Set(rows) : null)
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
