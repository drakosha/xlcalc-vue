// localSource: a sheet source over an xlcalc workbook whose session runs in a Web Worker pool or on the page's thread.
// The session holds and calculates the workbook, finds as Excel's Find does and returns windows, sheets and edits as Excel
// makes them; this side caches the columns of displayed values the display filter read, patched by changes, dropped on reload.
import { createInlineCalc, FUNCTION_NAMES, functionSignature, LOCALES, toLocal, type AsyncSession, type BookInput, type ExternalBook, type Calc, type CellChange, type CellPatch, type EditResult, type LayoutEdit, type Locale, type PatchFields, type RangeRef, type SheetExtent, type SheetLayout, type SheetView, type SplitPanes } from 'xlcalc'
import { cellAddress, linesKey, nameKey, outlineGroups, parseCell, parseRange, rangeAddress, sameName } from 'xlcalc/a1'
import type { ClearKind } from '../contract/workbook'
import type { CalculationMode, CellEntry, CellUpdate, EditOptions, EditOutcome, FindRequest, FormatPatch, LayoutChange, OutlineChange, ReplaceOutcome, SheetFilter, SheetInfo, SheetOp, SheetShow, SourceEvent, WindowRequest } from '../contract/source'
import { boxesUnder, findHits, replaceAll } from './find'
import { loadLocale, tipLocale } from './locale'
import { boxesOf, cellKey, failed, needsReload, outcomeOf, refusalOf, updateOf, type CellsBox } from './outcome'
import { throwIfAborted } from '../abort'
import { filterKey } from '../filter-key'
import { lineSpan, type LineAxis } from '../lines'
import { outlineOf, outlineOp as sharedOutlineOp } from '../memory-outline'
import { DEFAULT_MAX_CELLS, rowsFor } from '../window-size'
import { columnList, filterActive, filterColumns, visibleList } from '../sheet-query'
import { SheetValues } from './query'
import type { LocalSource, LocalSourceOptions } from './types'
import { manualCalculation } from './manual'
import { connections } from './connections'
import { protectionGuard, protectSheet } from './protect'
import { validationOf } from './validation'
import { autoFilterRanges } from './autofilter'
import { windowOf } from './window'
import { workbookMethods } from './workbook'

/** A display filter's rows as `visibleRows` last returned them and where the sheet's rows ended then; ascending (for search and totals) once requested. */
interface RowList {
  rows: number[]
  max: number
  asc?: number[]
}
/** How many filters' row lists a sheet keeps, for viewers that share the source with different filters. */
const LISTS_PER_SHEET = 8

/** The regular expressions' level the options ask for, registered with xlcalc before a calculation (a pool's threads take it too). */
async function loadRegex(level: LocalSourceOptions['regex']): Promise<void> {
  if (!level) return
  // the root's registerRegex through the module, so that ./local still links against an xlcalc before 0.66
  const [levels, root] = await Promise.all([import('xlcalc/regex'), import('xlcalc')])
  root.registerRegex(levels[level])
}

async function startCalc(opts: LocalSourceOptions): Promise<Calc> {
  const kind = typeof opts.calc === 'string' ? opts.calc : typeof Worker === 'function' ? 'pool' : 'inline'
  if (kind === 'inline') return createInlineCalc()
  const { createCalcPool } = await import('xlcalc/worker')
  return createCalcPool({ threads: 1, ...opts.pool })
}

/** The workbook's default font when it has none of its own: Calibri 11, in which xlcalc measures widths and `####`. */
const CALIBRI_11 = { name: 'Calibri', size: 11 }

/** A sheet of the workbook's layout with where its cells end, the workbook's default font and the range its AutoFilter acts on, as the component lists it. */
function info(sheet: SheetLayout, used: SheetExtent, font: { name: string; size: number }, autoFilter?: string): SheetInfo {
  const f = sheet.frozen
  const outline = outlineOf(outlineGroups(sheet), sheet.outline)
  return {
    name: sheet.name,
    maxRow: used.rows,
    maxCol: used.columns,
    ...(f ? { freeze: { rows: f.rows ?? 0, cols: f.columns ?? 0 } } : {}),
    ...(sheet.hidden ? { hidden: sheet.hidden } : {}),
    ...(outline ? { outline } : {}),
    ...(sheet.view?.zoom ? { zoom: sheet.view.zoom } : {}),
    ...(sheet.tabColor ? { tabColor: sheet.tabColor } : {}),
    ...(sheet.view?.gridlines === false ? { gridlines: false } : {}),
    ...(sheet.view?.headings === false ? { headings: false } : {}),
    ...(autoFilter ? { autoFilter: { range: autoFilter } } : {}),
    ...(sheet.protection ? { protection: sheet.protection } : {}),
    // frozen panes take the place of a split, as in Excel
    ...(sheet.view?.split && !f ? { split: sheet.view.split } : {}),
    font,
  }
}

/** Rows or columns as `setLayout` and the structure operations name them (`C`, `B:D`; `5`, `5:9`), from what the component sends. */
function linesOf(text: string, axis: LineAxis): string | null {
  const span = lineSpan(text, axis)
  return span && linesKey(span.first, span.last, axis)
}

/**
 * Excel's outline command as xlcalc's structure operation (its docs/outline.md), or why there is none, by the rule both
 * sources share (`outlineOp`); the sheet's groups are read only for a group's button.
 */
async function outlineOp(s: AsyncSession, change: OutlineChange): Promise<SheetOp | EditOutcome> {
  let layout: SheetLayout | undefined
  if ('level' in change && change.op !== 'showLevel') {
    layout = (await s.getLayout()).sheets.find((x) => sameName(x.name, change.sheet))
    if (!layout) return failed(`No sheet ${change.sheet}`)
  }
  const op = sharedOutlineOp(change, (axis) => (layout ? outlineGroups(layout)[axis] : []))
  return 'refused' in op ? { ok: false, reason: 'refused', message: op.refused } : ({ ...op, sheet: change.sheet } as SheetOp)
}

/** A source over `input`: the workbook is passed to the session as is (its cells are not copied), so do not change it afterwards. */
export function localSource(input: BookInput, opts: LocalSourceOptions = {}): LocalSource {
  const locale = input.locale ?? 'en-US'
  // Function descriptions, for Insert Function: loaded with the first tooltip, since a host that only calculates needs none of them.
  let describe: Promise<((fn: string, words: Locale | string) => string | null) | null> | undefined
  const maxCells = opts.maxCells ?? DEFAULT_MAX_CELLS
  const listeners = new Set<(event: SourceEvent) => void>()
  /** Displayed values of the sheets that the display filter or its funnel needed, by the sheet's name key. */
  const values = new Map<string, Promise<SheetValues>>()
  /** The rows of a sheet's display filters by filter, as `visibleRows` last returned them: windows under a filter use its list. */
  const lists = new Map<string, Map<string, RowList>>()
  /**
   * Excel's View > Show as set by `show`, by the sheet's name key: the session has no edit for it, since in Excel it is not
   * an undo step, so the source puts it on the sheets it lists and the workbook it returns, and follows renamed sheets.
   */
  const shows = new Map<string, SheetShow & { split?: SplitPanes | null }>()
  /** The settings of `shows` of deleted sheets, by the name key, for an undo that brings a sheet back. */
  const dropped = new Map<string, SheetShow & { split?: SplitPanes | null }>()
  const viewed = <T extends { name: string; view?: SheetView }>(sheet: T): T => {
    const own = shows.get(nameKey(sheet.name))
    if (!own) return sheet
    // View > Split set by `split`, null where it was removed
    const { split, ...shown } = own
    const view: SheetView = { ...sheet.view, ...shown }
    if (split) view.split = split
    else if (split === null) delete view.split
    return { ...sheet, view }
  }
  let calc: Calc | null = null
  let session: AsyncSession | null = null

  /** An answer arrived after its call: the session's `settleConnections` after the pending changes, once for the answers so far. */
  let taking = false
  function answered(): void {
    if (taking) return
    taking = true
    void serial(async () => {
      taking = false
      if (outside.closed) return
      const s = await open()
      const got = await s.settleConnections()
      if (!outside.closed) await settle(s, got, [])
    })().catch(() => {})
  }

  /** Excel's sheet protection, which the session does not enforce: each change to a protected sheet is checked first. */
  const guard = protectionGuard(locale, maxCells)
  /** The ranges the sheets' AutoFilters act on, which data typed under them grows: an edit that moves one sends `sheets`. */
  const filters = autoFilterRanges()

  /** The source's own changes in sequence: each starts from the state the previous one left. */
  let last: Promise<unknown> = Promise.resolve()
  function serial<A extends unknown[], R>(run: (...args: A) => Promise<R>): (...args: A) => Promise<R> {
    return (...args) => {
      const next = last.then(() => run(...args))
      last = next.catch(() => {})
      return next
    }
  }

  /** The host's answers that arrive after the call that asked for them, taken as they come (`answered`). */
  const outside = connections(opts.session, () => answered())

  const ready = (async () => {
    await Promise.all([loadLocale(locale), loadRegex(opts.regex)])
    calc = typeof opts.calc === 'object' ? opts.calc : await startCalc(opts)
    // a formula over a function the engine does not know keeps the value its file stores (or its cell holds), as Excel shows it until it calculates the cell
    session = await calc.createSession(input, { unknownFunctions: 'held', ...outside.session })
  })()
  ready.catch(() => {})
  const open = async (): Promise<AsyncSession> => {
    await ready
    return session!
  }

  function emit(event: SourceEvent): void {
    for (const l of listeners) l(event)
  }
  /** Excel's calculation option, which the session records and does not apply: in Manual the values shown wait for a calculation. */
  // told after the change's own events, which the caller sends once `take` returns
  const calcMode = manualCalculation(input, (calculation) => queueMicrotask(() => emit({ kind: 'calculation', calculation })))

  /** The workbook changed in many places: cached values are dropped and subscribers read the sheets and windows again. */
  function reload(): void {
    values.clear()
    lists.clear()
    emit({ kind: 'reload' })
  }
  /** Whether a conditional format of the workbook displays numbers with its own number format, which changes the cells' displayed text. */
  const rulesFormat = input.sheets.some((sh) => sh.conditionalFormats?.some((r) => 'style' in r && r.style?.numberFormat !== undefined))
  /**
   * Conditional formats may change how cells look anywhere their rules apply (the library's `repaint`), which cell updates
   * do not carry: subscribers re-read the windows over those ranges. A rule changes no value, so the cached values are kept
   * unless a rule formats numbers (re-reading a sheet's values is a pass over all its cells). The display filters' row lists
   * are kept: the component keeps its list until the filter is reapplied, as Excel keeps rows until Reapply.
   */
  function repaint(ranges: RangeRef[]): void {
    if (rulesFormat) values.clear()
    emit({ kind: 'repaint', ranges: ranges.map((x) => ({ sheet: x.sheet, range: x.range })) })
  }

  /** The session's changes as cell updates, with the edited cells' updates in place, and the cached values updated, one sheet at a time. */
  function take(changes: CellChange[], edited: Map<string, CellUpdate> = new Map()): CellUpdate[] {
    const out = new Map(edited)
    // Excel calculates a formula entered, typed, pasted or filled, with the cells its array fills
    const entered = new Set([...edited.keys(), ...changes.filter((c) => typeof c.formula === 'string').map((c) => cellKey(c.sheet, c.cell))])
    for (const ch of changes) {
      const key = cellKey(ch.sheet, ch.cell)
      if (!out.has(key) && !calcMode.hold(ch, entered)) out.set(key, { sheet: ch.sheet, cell: ch.cell, text: ch.text ?? '', value: ch.value, ...(ch.color ? { color: ch.color } : {}) })
    }
    const list = [...out.values()]
    calcMode.showing(list, changes)
    calcMode.report()
    patch(list)
    return list
  }

  /** The cached values of the sheets that cells updated, one sheet at a time. */
  function patch(list: CellUpdate[]): void {
    const bySheet = new Map<Promise<SheetValues>, CellUpdate[]>()
    for (const u of list) {
      const kept = values.get(nameKey(u.sheet))
      if (!kept) continue
      const mine = bySheet.get(kept)
      if (mine) mine.push(u)
      else bySheet.set(kept, [u])
    }
    for (const [kept, updates] of bySheet) {
      void kept.then((v) => {
        for (const u of updates) {
          const at = parseCell(u.cell)
          if (at) v.set(at.row, at.col, u.text, u.value)
        }
      })
    }
  }

  /** A sheet's displayed values: where its cells end, read once, and its columns as the filter needs them. */
  function sheetValues(s: AsyncSession, sheet: string): Promise<SheetValues> {
    const key = nameKey(sheet)
    let got = values.get(key)
    if (!got) {
      // in Manual the values shown, as the filter reads them in Excel
      got = s.getUsedRange(sheet).then((used) => new SheetValues(used.rows, async (ranges, signal) => calcMode.overlayCells(sheet, await s.getCells(sheet, ranges, signal ? { signal } : {}), ranges)))
      values.set(key, got)
      got.catch(() => values.delete(key))
    }
    return got
  }

  /** A sheet's values with the columns `cols` read, the read stopped by `signal`. */
  async function columnsOf(s: AsyncSession, sheet: string, cols: number[], signal?: AbortSignal): Promise<SheetValues> {
    const v = await sheetValues(s, sheet)
    await v.load(cols, signal)
    return v
  }

  /** The rows of a sheet under a filter: the list `visibleRows` returned for it, so windows match what the component holds. */
  async function rowsUnder(s: AsyncSession, sheet: string, filter: SheetFilter, fresh: boolean, signal?: AbortSignal): Promise<RowList> {
    const key = filterKey(filter)
    let byFilter = lists.get(nameKey(sheet))
    const kept = byFilter?.get(key)
    if (!fresh && kept) return kept
    const v = await columnsOf(s, sheet, filterColumns(filter), signal)
    const list: RowList = { rows: visibleList(v, filter, locale), max: v.maxRow }
    if (!byFilter) lists.set(nameKey(sheet), (byFilter = new Map()))
    // the newest last; the oldest list goes beyond the limit
    byFilter.delete(key)
    byFilter.set(key, list)
    if (byFilter.size > LISTS_PER_SHEET) byFilter.delete(byFilter.keys().next().value!)
    return list
  }

  /** The rows a search or totals read under the display filter; null without a filter or when it hides no row (a sort alone). */
  async function searchedRows(s: AsyncSession, sheet: string, filter: SheetFilter | undefined, signal?: AbortSignal): Promise<number[] | null> {
    if (!filterActive(filter)) return null
    const list = await rowsUnder(s, sheet, filter, false, signal)
    // the library takes the sheet's rows in ascending order; a filter that hides none searches as if there were none
    return list.rows.length < list.max ? (list.asc ??= [...list.rows].sort((a, b) => a - b)) : null
  }

  /** The edited cells in their current state, with their formulas as displayed and edit texts; `truncated` when a sheet's range had more than a window. */
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

  /** A completed change: its updates, with the edited cells' current state, sent to subscribers, then the formatting it may have changed; a reload where updates cannot carry it. */
  async function settle(s: AsyncSession, got: EditResult, boxes: CellsBox[]): Promise<Extract<EditOutcome, { ok: true }>> {
    const { updates, truncated } = await editedIn(s, boxes)
    const changes = take([...got.changes, ...(await calcMode.spillsOf(s, updates.values()))], updates)
    if (truncated || needsReload(got)) reload()
    else {
      if (changes.length) emit({ kind: 'changed', changes })
      if (got.repaint?.length) repaint(got.repaint)
      if (await filters.moved(s, changes)) emit({ kind: 'sheets' })
    }
    return { ok: true, changes }
  }

  /** A change of the workbook: after the source's pending changes, on the session, with a session exception turned into its outcome by `failure`. */
  function mutation<A extends unknown[], R>(run: (s: AsyncSession, ...args: A) => Promise<R>, failure: (e: unknown) => R): (...args: A) => Promise<R> {
    return serial(async (...args: A) => {
      const s = await open()
      try {
        await calcMode.capture(s)
        return await run(s, ...args)
      } catch (e) {
        return failure(e)
      } finally {
        // the change may have protected a sheet or changed a style, which the guard reads from the layout
        guard.forget()
      }
    })
  }
  /** An exception as the outcome of a change of typed input (the workbook rejects it) or of another change (it failed). */
  const refusedBy = (e: unknown) => outcomeOf(e, 'refused')
  const failedBy = (e: unknown) => outcomeOf(e, 'failed')

  /** A change the windows cannot carry: the cached values are updated, then the sheets and windows are read again. */
  function reloaded(changes: CellChange[]): EditOutcome {
    const out = take(changes)
    reload()
    return { ok: true, changes: out }
  }

  /**
   * Typed texts, each parsed as typed into its cell, in one step; only empty entries (Delete over a selection) clear the contents.
   * A value that violates a data validation rule of any style rejects the call with the rule's alert before anything is
   * written, as Excel's alert holds back the entry; the library picks the strictest rule a batch violates. With `accept`
   * (Yes on a warning, OK on an information alert) the texts are written without checking the rules again.
   */
  async function edit(s: AsyncSession, entries: CellEntry[], opts?: EditOptions): Promise<EditOutcome> {
    const cells: { sheet: string; cell: string; text: string | null }[] = []
    for (const e of entries) {
      const at = parseCell(e.cell)
      if (!at) return failed(`No cell ${e.sheet}!${e.cell}`)
      cells.push({ sheet: e.sheet, cell: cellAddress(at.row, at.col), text: e.text })
    }
    if (!cells.length) return { ok: true, changes: [] }
    const boxes = boxesOf(cells)
    if (cells.every((c) => c.text === null)) return settle(s, await s.clearCells(cells.map((c) => ({ sheet: c.sheet, range: c.cell })), 'contents'), boxes)
    const texts = cells.map((c) => ({ sheet: c.sheet, cell: c.cell, text: c.text ?? '' }))
    const got = await s.setCellTexts(texts, { refuseInvalid: !opts?.accept, ...(opts?.inPlace ? { inPlace: true } : {}) })
    return got.ok ? settle(s, got, boxes) : refusalOf(got)
  }

  /** Clipboard text as Excel pastes it at a cell, as a single undo step; the selection covers the filled range. */
  async function paste(s: AsyncSession, sheet: string, cell: string, text: string): Promise<EditOutcome> {
    const got = await s.pasteText(sheet, cell, text)
    if (!got.ok) return refusalOf(got)
    const out = await settle(s, got, [{ sheet, rect: parseRange(got.range)! }])
    return { ...out, select: { sheet, range: got.range } }
  }

  /**
   * An item chosen from a list's dropdown, written as Excel writes it: the item's `input` from the session (a cell's value
   * as is, a literal item parsed as typed in en-US), the cell's number format kept, no rule checked, a single undo step. The
   * list is read again, as Excel recalculates it, and an item whose text is no longer at its position is rejected.
   */
  async function pick(s: AsyncSession, sheet: string, cell: string, item: number, text: string): Promise<EditOutcome> {
    const at = parseCell(cell)
    if (!at) return failed(`No cell ${sheet}!${cell}`)
    const list = (await s.getValidation(sheet, cell))?.list
    const got = Array.isArray(list) ? list[item] : undefined
    if (!got || got.text !== text) return { ok: false, reason: 'conflict', message: `The list of ${sheet}!${cell} changed: open it again` }
    const addr = cellAddress(at.row, at.col)
    return settle(s, await s.setCells([{ sheet, cell: addr, input: got.input }]), boxesOf([{ sheet, cell: addr }]))
  }

  /** Ctrl+Enter: the text parsed once in the active cell and entered into every cell of the range, which stays selected; the active cell's rule rejects as in `edit`. */
  async function fill(s: AsyncSession, sheet: string, range: string, text: string, active: string, opts?: EditOptions): Promise<EditOutcome> {
    const rect = parseRange(range)
    if (!rect) return failed(`Not a range: ${range}`)
    const got = await s.fillCellText(sheet, range, text, { active, refuseInvalid: !opts?.accept, ...(opts?.inPlace ? { inPlace: true } : {}) })
    if (!got.ok) return refusalOf(got)
    const out = await settle(s, got, [{ sheet, rect }])
    return { ...out, select: { sheet, range: rangeAddress(rect) } }
  }

  /** The session's undo or redo of a whole change; the selection covers the cells it restored, on the first of their sheets. */
  async function step(s: AsyncSession, kind: 'undo' | 'redo'): Promise<EditOutcome> {
    const got = await s[kind]()
    if (!got) return { ok: true, changes: [] }
    // The step's changes keep the cached values current and the filters' row lists stay as the component holds them; a step that needs more reloads, which drops both.
    // TODO(xlcalc): undoing a structure step does not list the cells it empties (a copy onto empty cells), so the sheet is read again, as after the operation.
    if (got.call === 'editStructure') {
      calcMode.moved(got.moved)
      return reloaded(got.changes)
    }
    const boxes = boxesOf(got.cells ?? [])
    const out = await settle(s, got, boxes)
    const first = boxes[0]
    return first ? { ...out, select: { sheet: first.sheet, range: rangeAddress(first.rect) } } : out
  }

  /**
   * Runs a call that may rename, delete or bring back sheets, moving the settings of `shows` to a renamed sheet's new
   * name and keeping a deleted one's in `dropped`, which a step of undo or redo (`step`) that brings the sheet back restores.
   */
  async function following<R>(s: AsyncSession, run: () => Promise<R>, step = false): Promise<R> {
    if (!shows.size && !dropped.size) return run()
    const names = async () => new Set((await s.getLayout()).sheets.map((x) => nameKey(x.name)))
    const before = await names()
    const out = await run()
    const after = await names()
    const gone = [...before].filter((k) => !after.has(k))
    const came = [...after].filter((k) => !before.has(k))
    if (gone.length === 1 && came.length === 1) {
      const own = shows.get(gone[0]!)
      if (own) shows.set(came[0]!, own)
      shows.delete(gone[0]!)
      return out
    }
    for (const k of gone) {
      const own = shows.get(k)
      if (own) dropped.set(k, own)
      shows.delete(k)
    }
    for (const k of step ? came : []) {
      const own = dropped.get(k)
      if (own) shows.set(k, own)
      dropped.delete(k)
    }
    return out
  }

  /** A structure operation; the values waiting in Manual follow the addresses it moved (`calcMode.moved`). */
  async function structure(s: AsyncSession, op: SheetOp): Promise<EditOutcome> {
    const got = await s.editStructure(op)
    if (!got.ok) return { ok: false, reason: 'refused', code: got.code, message: got.message }
    calcMode.moved(got.moved)
    return reloaded(got.changes)
  }

  /** A structure operation once the sheet's protection allows it. */
  const guardedStructure = guard.structure(structure)

  /** Calculate Now, or Calculate Sheet: the values held in Manual shown, as an edit's changes are. */
  async function calculate(s: AsyncSession, sheet?: string): Promise<EditOutcome> {
    const changes = await calcMode.calculate(s, sheet)
    patch(changes)
    if (changes.length) emit({ kind: 'changed', changes })
    calcMode.report()
    return { ok: true, changes }
  }
  async function setCalculation(s: AsyncSession, mode: CalculationMode): Promise<EditOutcome> {
    const changes = await calcMode.setMode(s, mode)
    patch(changes)
    if (changes.length) emit({ kind: 'changed', changes })
    calcMode.report()
    return { ok: true, changes }
  }
  /** xlcalc's `setExternalBooks` (0.62.0): the links replaced, the workbook calculated again in one step of undo, its changes sent as an edit's. */
  async function externalBooks(s: AsyncSession, books: ExternalBook[]): Promise<EditOutcome> {
    if (typeof s.setExternalBooks !== 'function') return failed('Links to other workbooks need xlcalc 0.62.0 or later')
    return settle(s, await s.setExternalBooks(books), [])
  }

  /**
   * Excel's Refresh All (xlcalc 0.68's `refreshConnections`): the answers kept are dropped and the calls asked again, not a
   * step of undo; its changes are sent as an edit's, so in Manual the cells it calculated wait for a calculation.
   */
  async function refresh(s: AsyncSession, names?: string[]): Promise<EditOutcome> {
    return settle(s, await s.refreshConnections(names), [])
  }

  /** A change of what the sheet shows that is no change of the workbook: after the pending changes, outside undo. */
  const showing = <A extends unknown[]>(run: (s: AsyncSession, ...args: A) => Promise<EditOutcome>) => serial(async (...args: A) => run(await open(), ...args).catch(failedBy))

  async function setLayout(s: AsyncSession, change: LayoutChange): Promise<EditOutcome> {
    const shown = (size: number | undefined) => (change.hidden !== undefined ? { hidden: change.hidden ? true : null } : size !== undefined && size > 0 ? { hidden: null } : {})
    let edit: LayoutEdit
    if ('columns' in change) {
      const columns = linesOf(change.columns, 'columns')
      if (!columns) return failed(`Not columns: ${change.columns}`)
      // a manually set width is no longer the fitted one, which typing widens
      edit = change.fit ? { sheet: change.sheet, columns, fit: true } : { sheet: change.sheet, columns, set: { ...(change.width !== undefined ? { width: change.width, bestFit: null } : {}), ...shown(change.width) } }
    } else {
      const rows = linesOf(change.rows, 'rows')
      if (!rows) return failed(`Not rows: ${change.rows}`)
      edit = change.fit ? { sheet: change.sheet, rows, fit: true } : { sheet: change.sheet, rows, set: { ...(change.height !== undefined ? { height: change.height, bestFit: null } : {}), ...shown(change.height) } }
    }
    const got = await s.setLayout([edit])
    // The component already displays a size set with the mouse; it re-reads rows or columns that were hidden, unhidden or fitted.
    if (change.hidden !== undefined || ('fit' in change && change.fit)) return reloaded(got.changes)
    const changes = take(got.changes)
    if (changes.length) emit({ kind: 'changed', changes })
    return { ok: true, changes }
  }

  /** Formatting as Excel formats a selection: the style applied on top of each cell's displayed style, in one patch of the session. */
  async function format(s: AsyncSession, sheet: string, patches: FormatPatch[]): Promise<EditOutcome> {
    const cells: CellPatch[] = []
    for (const p of patches) {
      const set: PatchFields = {}
      if (p.style !== undefined) set.style = p.style
      if (p.numberFormat !== undefined) set.numberFormat = p.numberFormat
      if (Object.keys(set).length) cells.push({ sheet, range: p.range, set })
    }
    if (!cells.length) return { ok: true, changes: [] }
    const got = await s.patchCells(cells)
    // A number format changes how cells are displayed and what they open with, which windows carry: the sheet is read again.
    const changes = take(got.changes)
    if (cells.some((c) => c.set.numberFormat !== undefined)) {
      reload()
      return { ok: true, changes }
    }
    // A font that fits rows to it (xlcalc 0.59's `fitted`) changes row heights: the sheets and windows are read again, and the
    // filter's cached values stay, as a style changes no value.
    if (got.fitted?.length) {
      emit({ kind: 'reload' })
      return { ok: true, changes }
    }
    // A style alone changes no text or size, so the windows over the ranges are re-read in place.
    emit({ kind: 'repaint', ranges: [...patches.filter((p) => p.style !== undefined).map((p) => ({ sheet, range: p.range })), ...(got.repaint ?? []).map((x) => ({ sheet: x.sheet, range: x.range }))] })
    return { ok: true, changes }
  }

  /**
   * Excel's Replace All: the cells Find finds in formulas, whatever `lookIn` says (`./find`), each written as typed with its
   * matches replaced, in one `setCellTexts` and so a single undo step. As with Excel's paste, a replacement neither widens a
   * column nor checks data validation (not measured in Excel). A text the session rejects rejects the call, and nothing is written.
   */
  async function replace(s: AsyncSession, req: FindRequest, by: string): Promise<ReplaceOutcome> {
    const got = await replaceAll(s, req, by, await searchedRows(s, req.sheet, req.filter))
    const refused = got.refused?.length ? { refused: got.refused.map((r) => ({ sheet: r.sheet, cell: r.cell, message: r.message })) } : {}
    if (!got.replaced) return { ok: true, changes: [], replaced: 0, ...refused }
    // the written cells come back with their formulas and edit texts, read in windows of the cell limit
    const cells = got.changes.filter((c) => sameName(c.sheet, req.sheet)).map((c) => parseCell(c.cell)!)
    return { ...(await settle(s, got, boxesUnder(req.sheet, cells, maxCells))), replaced: got.replaced, ...refused }
  }

  /** Excel's outline commands as xlcalc's (`outlineOp`): a structure operation, a single undo step that recalculates SUBTOTAL and AGGREGATE; a rejection is the library's, in the workbook's language. */
  async function outline(s: AsyncSession, change: OutlineChange): Promise<EditOutcome> {
    const op = await outlineOp(s, change)
    return 'ok' in op ? op : structure(s, op)
  }

  /** Excel's Freeze Panes: the sheet's frozen rows and columns in one `setLayout` call, a single undo step (null or nothing frozen unfreezes); subscribers read the sheets again (`reload`), and the filter's cached values stay. */
  async function freeze(s: AsyncSession, sheet: string, panes: { rows: number; cols: number } | null): Promise<EditOutcome> {
    const frozen = panes && (panes.rows > 0 || panes.cols > 0) ? { ...(panes.rows > 0 ? { rows: panes.rows } : {}), ...(panes.cols > 0 ? { columns: panes.cols } : {}) } : null
    const got = await s.setLayout([{ sheet, frozen }])
    const changes = take(got.changes)
    emit({ kind: 'reload' })
    return { ok: true, changes }
  }

  /** Excel's Tab Color in one `setLayout` call, a single undo step (null: No Color); subscribers read the sheets again (`reload`). */
  async function tabColor(s: AsyncSession, sheet: string, color: string | null): Promise<EditOutcome> {
    const got = await s.setLayout([{ sheet, tabColor: color }])
    const changes = take(got.changes)
    emit({ kind: 'reload' })
    return { ok: true, changes }
  }

  /** Excel's View > Show: kept by the source (`shows`), not an undo step; subscribers read the sheets again (`reload`). A sheet the workbook does not have is refused. */
  async function show(s: AsyncSession, sheet: string, what: SheetShow): Promise<EditOutcome> {
    if (!(await s.getLayout()).sheets.some((x) => sameName(x.name, sheet))) return failed(`No sheet ${sheet}`)
    shows.set(nameKey(sheet), { ...shows.get(nameKey(sheet)), ...what })
    emit({ kind: 'reload' })
    return { ok: true, changes: [] }
  }

  /** Excel's View > Split: kept by the source with View > Show (`shows`), not an undo step; written with the sheet's view. */
  async function split(s: AsyncSession, sheet: string, panes: SplitPanes | null): Promise<EditOutcome> {
    if (!(await s.getLayout()).sheets.some((x) => sameName(x.name, sheet))) return failed(`No sheet ${sheet}`)
    shows.set(nameKey(sheet), { ...shows.get(nameKey(sheet)), split: panes })
    return { ok: true, changes: [] }
  }

  /** A patch of cells that changes how they look but no text (a hyperlink, a note): the windows over them are re-read in place. */
  function repainted(got: EditResult, ranges: { sheet: string; range: string }[]): EditOutcome {
    const changes = take(got.changes)
    if (changes.length) emit({ kind: 'changed', changes })
    emit({ kind: 'repaint', ranges: [...ranges, ...(got.repaint ?? []).map((x) => ({ sheet: x.sheet, range: x.range }))] })
    return { ok: true, changes }
  }
  /** A change whose values Manual cannot hold by address: they show, and nothing waits, before the sheets are read again. */
  function restructured(changes: CellChange[]): EditOutcome {
    calcMode.forget()
    return reloaded(changes)
  }
  const book = workbookMethods(locale, { repainted, reloaded, restructured }, opts.spelling)
  const reading = <A extends unknown[], R>(run: (s: AsyncSession, ...args: A) => Promise<R>) => async (...args: A): Promise<R> => run(await open(), ...args)

  const source: LocalSource = {
    maxCells,
    locale,
    ready,
    async sheets(signal) {
      const s = await open()
      const layout = await s.getLayout()
      const [used, ranges] = await Promise.all([Promise.all(layout.sheets.map((x) => s.getUsedRange(x.name))), filters.list(s, layout.sheets)])
      throwIfAborted(signal)
      // xlcalc 0.59 gives the workbook's default font; an older one, or a workbook without it, is Calibri 11
      const font = layout.font ?? CALIBRI_11
      return layout.sheets.map((x, i) => info(viewed(x), used[i]!, font, ranges[i]))
    },
    async window(req: WindowRequest, signal) {
      const s = await open()
      throwIfAborted(signal)
      const cols = Math.max(0, req.cols)
      const count = rowsFor(maxCells, req.rows, cols)
      const filtered = filterActive(req.filter)
      let rows: number[]
      if (filtered) rows = (await rowsUnder(s, req.sheet, req.filter!, false, signal)).rows.slice(req.row - 1, req.row - 1 + count)
      else rows = Array.from({ length: count }, (_, i) => req.row + i)
      // the viewport takes rows in ascending order; the window keeps the display order
      const asked = [...rows].sort((a, b) => a - b)
      const top = asked[0] ?? 1
      const range = `${cellAddress(top, req.col)}:${cellAddress(asked[asked.length - 1] ?? top, req.col + Math.max(1, cols) - 1)}`
      const v = await s.getViewport(req.sheet, range, { formulaLocal: true, editText: true, maxCells, ...(filtered || !rows.length ? { rows: asked } : {}) })
      throwIfAborted(signal)
      return calcMode.overlay(req.sheet, windowOf(v, { row: req.row, rows, col: req.col, cols, filtered, notice: opts.notice }), rows)
    },
    edit: mutation(guard.edit(edit), refusedBy),
    paste: mutation(guard.paste(paste), refusedBy),
    fill: mutation(guard.fill(fill), refusedBy),
    pick: mutation(guard.pick(pick), refusedBy),
    undo: mutation((s) => following(s, () => step(s, 'undo'), true), failedBy),
    redo: mutation((s) => following(s, () => step(s, 'redo'), true), failedBy),
    structure: mutation((s, op: SheetOp) => following(s, () => guardedStructure(s, op)), failedBy),
    layout: mutation(guard.layout(setLayout), failedBy),
    format: mutation(guard.format(format), refusedBy),
    replace: mutation(
      guard.replace(replace, (r): ReplaceOutcome => ({ ...r, replaced: 0 }), (s, req) => searchedRows(s, req.sheet, req.filter)),
      (e): ReplaceOutcome => ({ ...refusedBy(e), replaced: 0 }),
    ),
    outline: mutation(guard.outline(outline), failedBy),
    freeze: mutation(freeze, failedBy),
    tabColor: mutation(tabColor, failedBy),
    show: mutation(show, failedBy),
    split: mutation(split, failedBy),
    protect: mutation((s, sheet, protection) => protectSheet(s, sheet, protection, reloaded), failedBy),
    clear: mutation(guard.clear(async (s, sheet: string, ranges: string[], what: ClearKind) => book.clear(s, sheet, ranges, await guard.clearKind(s, sheet, what))), failedBy),
    link: mutation(guard.link(book.link), refusedBy),
    note: mutation(guard.note(book.note), refusedBy),
    names: reading(book.names),
    setNames: mutation(book.setNames, refusedBy),
    namesFromSelection: async (sheet, range, labels, signal) => {
      const s = await open()
      throwIfAborted(signal)
      return book.namesFromSelection(s, sheet, range, labels)
    },
    /** Excel's Workbook Statistics as the session counts it (xlcalc 0.67's `getWorkbookStatistics`). */
    async workbookStatistics(sheet, signal) {
      const s = await open()
      throwIfAborted(signal)
      const got = await s.getWorkbookStatistics(sheet)
      throwIfAborted(signal)
      return got
    },
    dependents: reading(book.dependents),
    rules: reading(book.rules),
    setRules: mutation(guard.setRules(book.setRules), refusedBy),
    table: mutation(guard.table(book.table), failedBy),
    setAutoFilter: mutation(guard.setAutoFilter(book.setAutoFilter), refusedBy),
    ...(opts.spelling ? { spelling: reading(book.spelling), learnWord: book.learnWord } : {}),
    /**
     * Excel's AutoSum as the session guesses it (xlcalc 0.59), its formula in the workbook's locale; null with an older library.
     * Excel guesses the same range for every function of its list, so the others are built here over the session's range.
     */
    async autoSum(sheet, cell, signal, fn = 'sum') {
      const s = await open()
      throwIfAborted(signal)
      if (typeof s.getAutoSum !== 'function') return null
      const got = await s.getAutoSum(sheet, cell)
      throwIfAborted(signal)
      if (fn === 'sum') return { range: got.range, formula: got.formulaLocal }
      const local = toLocal(`${fn.toUpperCase()}(${got.range ?? ''})`, locale)
      return local.ok ? { range: got.range, formula: local.formula } : null
    },
    autoSumFunctions: ['average', 'count', 'max', 'min'],
    /**
     * Excel's Go To Special as the session finds it (xlcalc 0.66's `getSpecialCells`). Excel's Last cell is the used range's
     * last row and column: those of the cells, spills included (`getUsedRange`), or further those of the rows and columns of
     * their own that the special cells' used range reaches (a height, a width, a style).
     */
    async specialCells(sheet, range, kind, signal) {
      const s = await open()
      throwIfAborted(signal)
      if (typeof s.getSpecialCells !== 'function') throw new Error('Go To Special needs xlcalc 0.66.0 or later')
      if (kind.type !== 'lastCell') return s.getSpecialCells(sheet, range, kind)
      const [used, ...found] = await Promise.all([s.getUsedRange(sheet), ...(['blanks', 'constants', 'formulas'] as const).map((type) => s.getSpecialCells(sheet, 'A1', { type }))])
      throwIfAborted(signal)
      let rows = used.rows
      let cols = used.columns
      for (const r of found.flat().map((a) => parseRange(a)!)) {
        rows = Math.max(rows, r.bottom)
        cols = Math.max(cols, r.right)
      }
      return [cellAddress(Math.max(1, rows), Math.max(1, cols))]
    },
    get history() {
      return session?.history ?? { undo: 0, redo: 0 }
    },
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    /** The session's copy of a range as Excel puts it on the clipboard (xlcalc 0.34): values formatted by their number formats, hidden rows and columns omitted; a range too large is rejected. */
    async copyText(sheet, range, opts) {
      const s = await open()
      throwIfAborted(opts?.signal)
      return s.copyText(sheet, range, { ...(opts?.trailingNewline === false ? { trailingNewline: false } : {}), ...(opts?.signal ? { signal: opts.signal } : {}) })
    },
    async validation(sheet, cell, signal) {
      const s = await open()
      throwIfAborted(signal)
      const got = validationOf(await s.getValidation(sheet, cell))
      throwIfAborted(signal)
      return got
    },
    /** Every function a formula can call, by its name in the workbook's locale (English where the locale has no name, as Excel ru-RU displays `ABS`), in the locale's alphabetical order. */
    async functionNames(signal) {
      await ready
      throwIfAborted(signal)
      const local = LOCALES.get(locale)?.functions ?? {}
      return FUNCTION_NAMES.map((name) => local[name] ?? name).sort(new Intl.Collator(locale).compare)
    },
    /** A function's argument tooltip and description in the workbook's locale (its localized names loaded on the first request), by its name in English or in the locale; a table lookup, and the component caches the result. */
    async functionSignature(name, signal) {
      await ready
      throwIfAborted(signal)
      const words = await tipLocale(locale)
      throwIfAborted(signal)
      const sig = functionSignature(name, words)
      describe ??= import('xlcalc/function-text').then((m) => m.functionDescription, () => null)
      const description = sig && (await describe)?.(sig.fn, words)
      return sig && description ? { ...sig, description } : sig
    },
    async explain(sheet, cell) {
      return (await (await open()).explain(sheet, cell)).text
    },
    /** Excel's Ctrl+arrow: the session's `getEnd`, measured as Excel's Range.End. */
    async end(sheet, cell, direction, signal) {
      const s = await open()
      throwIfAborted(signal)
      return s.getEnd(sheet, cell, direction)
    },
    /** Totals of a range: the session's summary without the cells, under a display filter over its rows only. */
    async stats(sheet, range, filter, signal) {
      const s = await open()
      throwIfAborted(signal)
      const rows = await searchedRows(s, sheet, filter, signal)
      try {
        const got = await s.summarizeRanges([{ sheet, range }], { ...(rows ? { rows } : {}), ...(signal ? { signal } : {}) })
        // an overflowing sum is #NUM!, with no number to display
        return { cells: got.cells, numbers: got.numbers, sum: typeof got.sum === 'number' ? got.sum : NaN, min: got.min, max: got.max }
      } catch (e) {
        throwIfAborted(signal)
        throw e
      }
    },
    async find(req, signal) {
      const s = await open()
      throwIfAborted(signal)
      const rows = await searchedRows(s, req.sheet, req.filter, signal)
      try {
        return await findHits(s, req, rows, signal)
      } catch (e) {
        // the library's rejection of an aborted call is the abort the component expects
        throwIfAborted(signal)
        throw e
      }
    },
    async visibleRows(sheet, filter, signal) {
      const s = await open()
      if (!filterActive(filter)) return { rows: null }
      throwIfAborted(signal)
      return { rows: (await rowsUnder(s, sheet, filter, true, signal)).rows }
    },
    async columnValues(sheet, col, filter, signal) {
      const s = await open()
      throwIfAborted(signal)
      return columnList(await columnsOf(s, sheet, [col, ...Object.keys(filter.cols ?? {}).map(Number)], signal), col, filter, locale)
    },
    async book() {
      const b = await (await open()).getBook()
      return calcMode.book(shows.size ? { ...b, sheets: b.sheets.map(viewed) } : b)
    },
    async snapshot() {
      const snap = await (await open()).getSnapshot()
      const book = shows.size ? { ...snap.book, sheets: snap.book.sheets.map(viewed) } : snap.book
      return { ...snap, book: calcMode.book(book) }
    },
    async calculation(signal) {
      await ready
      throwIfAborted(signal)
      return calcMode.state()
    },
    setExternalBooks: mutation(externalBooks, refusedBy),
    // Refresh All asks the host again: a source without `externals` has no calls to ask
    ...(opts.session?.externals ? { refreshConnections: mutation(refresh, failedBy) } : {}),
    calculate: showing(calculate),
    setCalculation: showing(setCalculation),
    async close() {
      outside.close()
      await ready.catch(() => {})
      await session?.close()
      if (calc && typeof opts.calc !== 'object') await calc.close()
    },
  }
  return source
}
