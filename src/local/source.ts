// localSource: a sheet source over an xlcalc book whose session runs in a Web Worker pool or in the page's thread.
// The session holds the book, computes it, finds as Excel's Find and answers windows, sheets and edits as Excel makes
// them; this side keeps the sheets' shown values for the display filter, patched by the changes and dropped on a reload.
import { createInlineCalc, FUNCTION_NAMES, LOCALES, type AsyncSession, type BookInput, type Calc, type CellChange, type CellPatch, type EditResult, type LayoutEdit, type LineSpan, type PatchFields, type RangeRef, type SheetExtent, type SheetLayout } from 'xlcalc'
import { cellAddress, colLetters, linesKey, nameKey, outlineGroups, parseCell, parseRange, rangeAddress, sameName } from 'xlcalc/a1'
import type { CellEntry, CellUpdate, EditOptions, EditOutcome, FindRequest, FormatPatch, LayoutChange, OutlineChange, ReplaceOutcome, SheetFilter, SheetInfo, SheetOp, SheetOutline, SourceEvent, WindowRequest } from '../source'
import { boxesUnder, findHits, replaceAll } from './find'
import { loadLocale } from './locale'
import { boxesOf, cellKey, failed, needsReload, outcomeOf, refusalOf, updateOf, type CellsBox } from './outcome'
import { filterKey } from '../filter-key'
import { columnValuesOf, filterActive, SheetValues, statsOf, visibleList } from './query'
import type { LocalSource, LocalSourceOptions } from './types'
import { alertRefusalOf, validationOf } from './validation'
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
  const outline = outlineOf(sheet)
  return {
    name: sheet.name,
    maxRow: used.rows,
    maxCol: used.columns,
    ...(f ? { freeze: { rows: f.rows ?? 0, cols: f.columns ?? 0 } } : {}),
    ...(sheet.hidden ? { hidden: sheet.hidden } : {}),
    ...(outline ? { outline } : {}),
  }
}

/** A sheet's groups as xlcalc draws them (`outlineGroups`, Excel's marks), in the contract's terms; none without groups. */
function outlineOf(sheet: SheetLayout): SheetOutline | undefined {
  const { rows, columns } = outlineGroups(sheet)
  if (!rows.length && !columns.length) return undefined
  const groups = (list: typeof rows) => list.map((g) => ({ from: g.first, to: g.last, level: g.level, collapsed: g.collapsed }))
  return { rows: groups(rows), columns: groups(columns), summaryBelow: sheet.outline?.summaryBelow ?? true, summaryRight: sheet.outline?.summaryRight ?? true }
}

/** The lines of one axis the component sends (`C`, `C:C`, `B:D`; `3:4`), or null. */
function spanOf(text: string, axis: 'columns' | 'rows'): LineSpan | null {
  const got = parseRange(text.includes(':') ? text : `${text}:${text}`)
  if (got?.kind !== axis) return null
  return axis === 'columns' ? { first: got.left, last: got.right } : { first: got.top, last: got.bottom }
}

/** Lines as `setLayout` and the structure operations name them (`C`, `B:D`; `5`, `5:9`) from what the component sends. */
function linesOf(text: string, axis: 'columns' | 'rows'): string | null {
  const span = spanOf(text, axis)
  return span && linesKey(span.first, span.last, axis)
}

/**
 * Excel's outline command as xlcalc's structure operation (its docs/outline.md), or why there is none. A group's button is
 * Show or Hide Detail at its summary line, the group found among the sheet's groups by its lines and level; a group at the
 * sheet's edge has no summary line and, as in Excel, no button.
 */
async function outlineOp(s: AsyncSession, change: OutlineChange): Promise<SheetOp | EditOutcome> {
  const { sheet } = change
  const refused = (message: string): EditOutcome => ({ ok: false, reason: 'refused', message })
  if (change.op === 'showLevel') return change.axis === 'rows' ? { op: 'showLevels', sheet, rows: change.level } : { op: 'showLevels', sheet, columns: change.level }
  const [axis, text] = 'rows' in change ? (['rows', change.rows] as const) : (['columns', change.columns] as const)
  const span = spanOf(text, axis)
  if (!span) return refused(`Not ${axis}: ${text}`)
  if (!('level' in change)) {
    const lines = linesKey(span.first, span.last, axis)
    return axis === 'rows' ? { op: change.op, sheet, rows: lines } : { op: change.op, sheet, columns: lines }
  }
  const layout = (await s.getLayout()).sheets.find((x) => sameName(x.name, sheet))
  if (!layout) return failed(`No sheet ${sheet}`)
  const g = outlineGroups(layout)[axis].find((x) => x.first === span.first && x.last === span.last && x.level === change.level)
  if (!g) return refused(`No group of level ${change.level} over the ${axis} ${text}`)
  if (g.summary === null) return refused(`The group of ${axis} ${text} has no summary ${axis === 'rows' ? 'row' : 'column'}, past the sheet's edge: Excel draws no button for it`)
  const op = change.op === 'collapse' ? 'hideDetail' : 'showDetail'
  return axis === 'rows' ? { op, sheet, row: g.summary } : { op, sheet, column: colLetters(g.summary) }
}

/** A source over `input`: the book goes to the session as it is (its cells are not copied), so do not change it afterwards. */
export function localSource(input: BookInput, opts: LocalSourceOptions = {}): LocalSource {
  const locale = input.locale ?? 'en-US'
  // Text sorts in the book's locale once it is loaded; an unknown locale fails the opening.
  let collator: Intl.Collator
  const maxCells = opts.maxCells ?? MAX_CELLS
  const listeners = new Set<(event: SourceEvent) => void>()
  /** Shown values of sheets that the display filter or its funnel needed, by the sheet's name key. */
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
      if (!at) continue
      void values.get(nameKey(u.sheet))?.then((v) => v.set(at.row, at.col, u.text, u.value))
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

  /** The rows a search under the display filter looks in; null without a filter or when it hides no row (a sort alone). */
  async function searchedRows(s: AsyncSession, sheet: string, filter: SheetFilter | undefined): Promise<number[] | null> {
    if (!filterActive(filter)) return null
    const [rows, used] = await Promise.all([rowsUnder(s, sheet, filter, false), s.getUsedRange(sheet)])
    // the library takes the rows on the sheet in ascending order; a filter that hides none searches as without it
    return rows.length < used.rows ? [...rows].sort((a, b) => a - b) : null
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

  /**
   * An item picked from a list's dropdown, as Excel writes it: the item's `input` the session gives (a cell's value as it is,
   * a literal item read as typed in en-US), the cell's format kept, no rule asked, one step of the undo. The list is read
   * again, as Excel computes it, and an item whose text is no longer at its place is refused.
   */
  async function pick(sheet: string, cell: string, item: number, text: string): Promise<EditOutcome> {
    const s = await open()
    const at = parseCell(cell)
    if (!at) return failed(`No cell ${sheet}!${cell}`)
    try {
      const list = (await s.getValidation(sheet, cell))?.list
      const got = Array.isArray(list) ? list[item] : undefined
      if (!got || got.text !== text) return { ok: false, reason: 'conflict', message: `The list of ${sheet}!${cell} changed: open it again` }
      const addr = cellAddress(at.row, at.col)
      return await settle(s, await s.setCells([{ sheet, cell: addr, input: got.input }]), boxesOf([{ sheet, cell: addr }]))
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
   * Excel's Replace All: the cells Find finds in formulas, whatever `lookIn` says (`./find`), each written as typed with its
   * matches replaced, in one `setCellTexts` and so one step of undo. As Excel's paste does, a replacement neither widens a
   * column nor checks data validation (not measured in Excel). A text the session refuses refuses the call, nothing written.
   */
  async function replace(req: FindRequest, by: string): Promise<ReplaceOutcome> {
    const s = await open()
    try {
      const got = await replaceAll(s, req, by, await searchedRows(s, req.sheet, req.filter))
      const refused = got.refused?.length ? { refused: got.refused.map((r) => ({ sheet: r.sheet, cell: r.cell, message: r.message })) } : {}
      if (!got.replaced) return { ok: true, changes: [], replaced: 0, ...refused }
      // the written cells come back with their formulas and edit texts, read in windows of the cell limit
      const cells = got.changes.filter((c) => sameName(c.sheet, req.sheet)).map((c) => parseCell(c.cell)!)
      return { ...(await settle(s, got, boxesUnder(req.sheet, cells, maxCells))), replaced: got.replaced, ...refused }
    } catch (e) {
      return { ...outcomeOf(e, 'refused'), replaced: 0 }
    }
  }


  /** Excel's outline commands as xlcalc's (`outlineOp`): a structure operation, one step of the undo with SUBTOTAL and AGGREGATE recomputed; a refusal is the library's, in the book's language. */
  async function outline(change: OutlineChange): Promise<EditOutcome> {
    const s = await open()
    try {
      const op = await outlineOp(s, change)
      return 'ok' in op ? op : await structure(op)
    } catch (e) {
      return outcomeOf(e, 'failed')
    }
  }

  /** Excel's Freeze Panes: the sheet's frozen rows and columns in one `setLayout` call, one step of the undo (null or nothing frozen unfreezes); subscribers get `reload`. */
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
    pick: serial(pick),
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
    /** Excel's Control and an arrow: the session's `getEnd`, measured as Excel's Range.End. */
    async end(sheet, cell, direction, signal) {
      const s = await open()
      check(signal)
      return s.getEnd(sheet, cell, direction)
    },
    async stats(sheet, range, filter, signal) {
      const s = await open()
      const [views, rows] = await Promise.all([s.getCells(sheet, range), filterActive(filter) ? rowsUnder(s, sheet, filter, false) : null])
      check(signal)
      return statsOf(views, rows ? new Set(rows) : null)
    },
    async find(req, signal) {
      const s = await open()
      const rows = await searchedRows(s, req.sheet, req.filter)
      check(signal)
      try {
        return await findHits(s, req, rows, signal)
      } catch (e) {
        // the library's refusal of an aborted call is the abort the component expects
        check(signal)
        throw e
      }
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
