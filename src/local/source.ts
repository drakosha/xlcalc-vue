// localSource: a sheet source over an xlcalc book whose session runs in a Web Worker pool or in the page's thread.
// The session holds the book, computes it, finds as Excel's Find and answers windows, sheets and edits as Excel makes
// them; this side keeps the columns of shown values the display filter read, patched by the changes and dropped on a reload.
import { createInlineCalc, FUNCTION_NAMES, functionSignature, LOCALES, type AsyncSession, type BookInput, type Calc, type CellChange, type CellPatch, type EditResult, type LayoutEdit, type Locale, type PatchFields, type RangeRef, type SheetExtent, type SheetLayout } from 'xlcalc'
import { cellAddress, linesKey, nameKey, outlineGroups, parseCell, parseRange, rangeAddress, sameName } from 'xlcalc/a1'
import type { CellEntry, CellUpdate, EditOptions, EditOutcome, FindRequest, FormatPatch, LayoutChange, OutlineChange, ReplaceOutcome, SheetFilter, SheetInfo, SheetOp, SourceEvent, WindowRequest } from '../source'
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
import { validationOf } from './validation'
import { windowOf } from './window'

/** A display filter's rows as `visibleRows` last gave them and where the sheet's rows ended then; ascending (a search's and totals') once asked. */
interface RowList {
  key: string
  rows: number[]
  max: number
  asc?: number[]
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
  const outline = outlineOf(outlineGroups(sheet), sheet.outline)
  return {
    name: sheet.name,
    maxRow: used.rows,
    maxCol: used.columns,
    ...(f ? { freeze: { rows: f.rows ?? 0, cols: f.columns ?? 0 } } : {}),
    ...(sheet.hidden ? { hidden: sheet.hidden } : {}),
    ...(outline ? { outline } : {}),
  }
}

/** Lines as `setLayout` and the structure operations name them (`C`, `B:D`; `5`, `5:9`) from what the component sends. */
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
  // What each function does, for Insert Function: loaded with the first tip, as a host that only computes needs none of it.
  let describe: Promise<((fn: string, words: Locale | string) => string | null) | null> | undefined
  const maxCells = opts.maxCells ?? DEFAULT_MAX_CELLS
  const listeners = new Set<(event: SourceEvent) => void>()
  /** Shown values of sheets that the display filter or its funnel needed, by the sheet's name key. */
  const values = new Map<string, Promise<SheetValues>>()
  /** The rows of a sheet's display filter, as last given by `visibleRows`: windows under the filter read them. */
  const lists = new Map<string, RowList>()
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

  /** The session's changes as cell updates, the edited cells' in their place, and the kept values brought up to date, a sheet's at once. */
  function take(changes: CellChange[], edited: Map<string, CellUpdate> = new Map()): CellUpdate[] {
    const out = new Map(edited)
    for (const ch of changes) {
      const key = cellKey(ch.sheet, ch.cell)
      if (!out.has(key)) out.set(key, { sheet: ch.sheet, cell: ch.cell, text: ch.text ?? '', value: ch.value, ...(ch.color ? { color: ch.color } : {}) })
    }
    const list = [...out.values()]
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
    return list
  }

  /** A sheet's shown values: where its cells end read once, its columns as the filter needs them. */
  function sheetValues(s: AsyncSession, sheet: string): Promise<SheetValues> {
    const key = nameKey(sheet)
    let got = values.get(key)
    if (!got) {
      got = s.getUsedRange(sheet).then((used) => new SheetValues(used.rows, (ranges, signal) => s.getCells(sheet, ranges, signal ? { signal } : {})))
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

  /** The rows of a sheet under a filter: the list `visibleRows` gave for it, so windows match what the component holds. */
  async function rowsUnder(s: AsyncSession, sheet: string, filter: SheetFilter, fresh: boolean, signal?: AbortSignal): Promise<RowList> {
    const key = filterKey(filter)
    const kept = lists.get(nameKey(sheet))
    if (!fresh && kept?.key === key) return kept
    const v = await columnsOf(s, sheet, filterColumns(filter), signal)
    const list: RowList = { key, rows: visibleList(v, filter, locale), max: v.maxRow }
    lists.set(nameKey(sheet), list)
    return list
  }

  /** The rows a search or totals under the display filter read; null without a filter or when it hides no row (a sort alone). */
  async function searchedRows(s: AsyncSession, sheet: string, filter: SheetFilter | undefined, signal?: AbortSignal): Promise<number[] | null> {
    if (!filterActive(filter)) return null
    const list = await rowsUnder(s, sheet, filter, false, signal)
    // the library takes the rows on the sheet in ascending order; a filter that hides none searches as without it
    return list.rows.length < list.max ? (list.asc ??= [...list.rows].sort((a, b) => a - b)) : null
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
  async function settle(s: AsyncSession, got: EditResult, boxes: CellsBox[]): Promise<Extract<EditOutcome, { ok: true }>> {
    const { updates, truncated } = await editedIn(s, boxes)
    const changes = take(got.changes, updates)
    if (truncated || needsReload(got)) reload()
    else {
      if (changes.length) emit({ kind: 'changed', changes })
      if (got.repaint?.length) repaint(got.repaint)
    }
    return { ok: true, changes }
  }

  /** A change of the book: after the source's changes before it, on the session, an exception of the session as its outcome by `failure`. */
  function mutation<A extends unknown[], R>(run: (s: AsyncSession, ...args: A) => Promise<R>, failure: (e: unknown) => R): (...args: A) => Promise<R> {
    return serial(async (...args: A) => {
      const s = await open()
      try {
        return await run(s, ...args)
      } catch (e) {
        return failure(e)
      }
    })
  }
  /** An exception as the outcome of a change of typed input (the book refuses it) or of another change (it failed). */
  const refusedBy = (e: unknown) => outcomeOf(e, 'refused')
  const failedBy = (e: unknown) => outcomeOf(e, 'failed')

  /** A change the windows cannot carry: the kept values brought up to date, then the sheets and windows read again. */
  function reloaded(changes: CellChange[]): EditOutcome {
    const out = take(changes)
    reload()
    return { ok: true, changes: out }
  }

  /**
   * Typed texts, each read as typed into its cell, in one step; only empty cells (Delete over a selection) clear the content.
   * A value a rule of data validation does not take, of any style, refuses the call with the rule's alert before anything is
   * written, as Excel's alert holds the entry; the library picks the strictest rule a batch breaks. With `accept` (a
   * warning's Yes, an information's OK) the texts are written without asking the rules again.
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

  /** Clipboard text as Excel pastes it at a cell, one step of the undo; the selection goes over what it filled. */
  async function paste(s: AsyncSession, sheet: string, cell: string, text: string): Promise<EditOutcome> {
    const got = await s.pasteText(sheet, cell, text)
    if (!got.ok) return refusalOf(got)
    const out = await settle(s, got, [{ sheet, rect: parseRange(got.range)! }])
    return { ...out, select: { sheet, range: got.range } }
  }

  /**
   * An item picked from a list's dropdown, as Excel writes it: the item's `input` the session gives (a cell's value as it is,
   * a literal item read as typed in en-US), the cell's format kept, no rule asked, one step of the undo. The list is read
   * again, as Excel computes it, and an item whose text is no longer at its place is refused.
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

  /** Control+Return: the text read once in the active cell and entered into every cell of the range, which stays selected; the active cell's rule refuses as in `edit`. */
  async function fill(s: AsyncSession, sheet: string, range: string, text: string, active: string, opts?: EditOptions): Promise<EditOutcome> {
    const rect = parseRange(range)
    if (!rect) return failed(`Not a range: ${range}`)
    const got = await s.fillCellText(sheet, range, text, { active, refuseInvalid: !opts?.accept, ...(opts?.inPlace ? { inPlace: true } : {}) })
    if (!got.ok) return refusalOf(got)
    const out = await settle(s, got, [{ sheet, rect }])
    return { ...out, select: { sheet, range: rangeAddress(rect) } }
  }

  /** The session's undo or redo, whole; the selection goes over the cells it took back, on the first of their sheets. */
  async function step(s: AsyncSession, kind: 'undo' | 'redo'): Promise<EditOutcome> {
    const got = await s[kind]()
    if (!got) return { ok: true, changes: [] }
    // The step's changes keep the read values current and the filters' lists stay as the component holds them; a step that needs more reloads, which drops both.
    // TODO(xlcalc): a structure step's undo does not list the cells it empties (a copy onto empty cells), so the sheet is read again, as after the operation.
    if (got.call === 'editStructure') return reloaded(got.changes)
    const boxes = boxesOf(got.cells ?? [])
    const out = await settle(s, got, boxes)
    const first = boxes[0]
    return first ? { ...out, select: { sheet: first.sheet, range: rangeAddress(first.rect) } } : out
  }

  async function structure(s: AsyncSession, op: SheetOp): Promise<EditOutcome> {
    const got = await s.editStructure(op)
    return got.ok ? reloaded(got.changes) : { ok: false, reason: 'refused', message: got.message }
  }

  async function setLayout(s: AsyncSession, change: LayoutChange): Promise<EditOutcome> {
    const shown = (size: number | undefined) => (change.hidden !== undefined ? { hidden: change.hidden ? true : null } : size !== undefined && size > 0 ? { hidden: null } : {})
    let edit: LayoutEdit
    if ('columns' in change) {
      const columns = linesOf(change.columns, 'columns')
      if (!columns) return failed(`Not columns: ${change.columns}`)
      // a width set by hand is no longer the fitted one, which typing widens
      edit = change.fit ? { sheet: change.sheet, columns, fit: true } : { sheet: change.sheet, columns, set: { ...(change.width !== undefined ? { width: change.width, bestFit: null } : {}), ...shown(change.width) } }
    } else {
      const rows = linesOf(change.rows, 'rows')
      if (!rows) return failed(`Not rows: ${change.rows}`)
      edit = change.fit ? { sheet: change.sheet, rows, fit: true } : { sheet: change.sheet, rows, set: { ...(change.height !== undefined ? { height: change.height, bestFit: null } : {}), ...shown(change.height) } }
    }
    const got = await s.setLayout([edit])
    // A size set with the mouse the component already shows; lines hidden, shown or fitted it reads again.
    if (change.hidden !== undefined || ('fit' in change && change.fit)) return reloaded(got.changes)
    const changes = take(got.changes)
    if (changes.length) emit({ kind: 'changed', changes })
    return { ok: true, changes }
  }

  /** Formatting as Excel formats a selection: the style over the one each cell shows, in one patch of the session. */
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
    // A number format changes what cells show and open with, which windows carry: the sheet is read again. A look alone changes no
    // text or size, so the windows over the ranges are read again in place and the values the filter read stay.
    if (cells.some((c) => c.set.numberFormat !== undefined)) return reloaded(got.changes)
    const changes = take(got.changes)
    emit({ kind: 'repaint', ranges: [...patches.filter((p) => p.style !== undefined).map((p) => ({ sheet, range: p.range })), ...(got.repaint ?? []).map((x) => ({ sheet: x.sheet, range: x.range }))] })
    return { ok: true, changes }
  }

  /**
   * Excel's Replace All: the cells Find finds in formulas, whatever `lookIn` says (`./find`), each written as typed with its
   * matches replaced, in one `setCellTexts` and so one step of undo. As Excel's paste does, a replacement neither widens a
   * column nor checks data validation (not measured in Excel). A text the session refuses refuses the call, nothing written.
   */
  async function replace(s: AsyncSession, req: FindRequest, by: string): Promise<ReplaceOutcome> {
    const got = await replaceAll(s, req, by, await searchedRows(s, req.sheet, req.filter))
    const refused = got.refused?.length ? { refused: got.refused.map((r) => ({ sheet: r.sheet, cell: r.cell, message: r.message })) } : {}
    if (!got.replaced) return { ok: true, changes: [], replaced: 0, ...refused }
    // the written cells come back with their formulas and edit texts, read in windows of the cell limit
    const cells = got.changes.filter((c) => sameName(c.sheet, req.sheet)).map((c) => parseCell(c.cell)!)
    return { ...(await settle(s, got, boxesUnder(req.sheet, cells, maxCells))), replaced: got.replaced, ...refused }
  }

  /** Excel's outline commands as xlcalc's (`outlineOp`): a structure operation, one step of the undo with SUBTOTAL and AGGREGATE recomputed; a refusal is the library's, in the book's language. */
  async function outline(s: AsyncSession, change: OutlineChange): Promise<EditOutcome> {
    const op = await outlineOp(s, change)
    return 'ok' in op ? op : structure(s, op)
  }

  /** Excel's Freeze Panes: the sheet's frozen rows and columns in one `setLayout` call, one step of the undo (null or nothing frozen unfreezes); subscribers read the sheets again (`reload`), the values the filter read stay. */
  async function freeze(s: AsyncSession, sheet: string, panes: { rows: number; cols: number } | null): Promise<EditOutcome> {
    const frozen = panes && (panes.rows > 0 || panes.cols > 0) ? { ...(panes.rows > 0 ? { rows: panes.rows } : {}), ...(panes.cols > 0 ? { columns: panes.cols } : {}) } : null
    const got = await s.setLayout([{ sheet, frozen }])
    const changes = take(got.changes)
    emit({ kind: 'reload' })
    return { ok: true, changes }
  }

  const source: LocalSource = {
    maxCells,
    locale,
    ready,
    async sheets(signal) {
      const s = await open()
      const { sheets } = await s.getLayout()
      const used = await Promise.all(sheets.map((x) => s.getUsedRange(x.name)))
      throwIfAborted(signal)
      return sheets.map((x, i) => info(x, used[i]!))
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
      return windowOf(v, { row: req.row, rows, col: req.col, cols, filtered, notice: opts.notice })
    },
    edit: mutation(edit, refusedBy),
    paste: mutation(paste, refusedBy),
    fill: mutation(fill, refusedBy),
    pick: mutation(pick, refusedBy),
    undo: mutation((s) => step(s, 'undo'), failedBy),
    redo: mutation((s) => step(s, 'redo'), failedBy),
    structure: mutation(structure, failedBy),
    layout: mutation(setLayout, failedBy),
    format: mutation(format, refusedBy),
    replace: mutation(replace, (e): ReplaceOutcome => ({ ...refusedBy(e), replaced: 0 })),
    outline: mutation(outline, failedBy),
    freeze: mutation(freeze, failedBy),
    get history() {
      return session?.history ?? { undo: 0, redo: 0 }
    },
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    /** The session's copy of a range as Excel puts it on the clipboard (xlcalc 0.34): values by their formats, hidden lines left out; a range too large is refused. */
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
    /** Every function a formula can call, by its name in the book's locale (English where the locale names none, as Excel ru-RU shows `ABS`), in the locale's alphabetical order. */
    async functionNames(signal) {
      await ready
      throwIfAborted(signal)
      const local = LOCALES.get(locale)?.functions ?? {}
      return FUNCTION_NAMES.map((name) => local[name] ?? name).sort(new Intl.Collator(locale).compare)
    },
    /** A function's tip and description in the book's locale (its words loaded on the first ask), by its name in English or the locale; a table lookup, the component keeps what it got. */
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
    /** Excel's Control and an arrow: the session's `getEnd`, measured as Excel's Range.End. */
    async end(sheet, cell, direction, signal) {
      const s = await open()
      throwIfAborted(signal)
      return s.getEnd(sheet, cell, direction)
    },
    /** Totals of a range: the session's summary without the cells, under a display filter over its rows. */
    async stats(sheet, range, filter, signal) {
      const s = await open()
      throwIfAborted(signal)
      const rows = await searchedRows(s, sheet, filter, signal)
      try {
        const got = await s.summarizeRanges([{ sheet, range }], { ...(rows ? { rows } : {}), ...(signal ? { signal } : {}) })
        // an overflowed sum is #NUM!, no number to show
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
        // the library's refusal of an aborted call is the abort the component expects
        throwIfAborted(signal)
        throw e
      }
    },
    async visibleRows(sheet, filter, signal) {
      const s = await open()
      if (!filterActive(filter)) {
        lists.delete(nameKey(sheet))
        return { rows: null }
      }
      throwIfAborted(signal)
      return { rows: (await rowsUnder(s, sheet, filter, true, signal)).rows }
    },
    async columnValues(sheet, col, filter, signal) {
      const s = await open()
      throwIfAborted(signal)
      return columnList(await columnsOf(s, sheet, [col, ...Object.keys(filter.cols ?? {}).map(Number)], signal), col, filter, locale)
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
