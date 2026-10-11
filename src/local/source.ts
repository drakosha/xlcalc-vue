// localSource: a sheet source over an xlcalc book whose session runs in a Web Worker pool or in the page's thread.
// The session holds the book, computes it and answers windows, sheets and edits as Excel makes them; this side keeps
// the sheets' shown values for search and the display filter, patched by the changes and dropped on a reload.
import { createInlineCalc, type AsyncSession, type BookInput, type Calc, type CellChange, type CellPatch, type EditResult, type PatchFields, type SheetExtent, type SheetLayout } from 'xlcalc'
import { cellAddress, linesKey, nameKey, parseCell, parseRange, rangeAddress } from 'xlcalc/a1'
import type { CellEntry, CellUpdate, EditOutcome, FormatPatch, LayoutChange, SheetFilter, SheetInfo, SheetOp, SourceEvent, WindowRequest } from '../source'
import { loadLocale } from './locale'
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
  return {
    name: sheet.name,
    maxRow: used.rows,
    maxCol: used.columns,
    ...(f ? { freeze: { rows: f.rows ?? 0, cols: f.columns ?? 0 } } : {}),
    ...(sheet.hidden ? { hidden: true } : {}),
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
  // TODO(xlcalc): a call a stop rule refused is taken back by undo() and stays on the redo stack, where Excel has nothing; this hides that one step.
  let refusedRedo = false

  /** The source's own changes one after another: each reads what the one before left. */
  let last: Promise<unknown> = Promise.resolve()
  function serial<A extends unknown[], R>(run: (...args: A) => Promise<R>): (...args: A) => Promise<R> {
    return (...args) => {
      const next = last.then(() => run(...args))
      last = next.catch(() => {})
      return next
    }
  }
  /** A change of the book through the source: a new step forgets what redo could do, the refused step too. */
  function change<A extends unknown[]>(run: (...args: A) => Promise<EditOutcome>): (...args: A) => Promise<EditOutcome> {
    return serial(async (...args: A) => {
      const before = session?.revision
      const out = await run(...args)
      if (out.ok && session && session.revision !== before) refusedRedo = false
      return out
    })
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

  /** A change made: its updates, the edited cells' as they are now, told to subscribers; a reload where updates cannot carry it. */
  async function settle(s: AsyncSession, got: EditResult, boxes: CellsBox[]): Promise<EditOutcome> {
    const { updates, truncated } = await editedIn(s, boxes)
    const changes = take(got.changes, updates)
    if (truncated || needsReload(got)) reload()
    else if (changes.length) emit({ kind: 'changed', changes })
    return { ok: true, changes }
  }

  /** A value a stop rule of data validation does not take: the call is taken back whole, as Excel takes it back. */
  async function refuseStop(s: AsyncSession, alert: { sheet: string; cell: string; message: string }): Promise<EditOutcome> {
    await s.undo()
    refusedRedo = true
    return { ok: false, reason: 'refused', message: alert.message, cells: [{ sheet: alert.sheet, cell: alert.cell }] }
  }

  /** Typed texts, each read as typed into its cell, in one step; only empty cells (Delete over a selection) clear the content. */
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
      const got = await s.setCellTexts(cells.map((c) => ({ sheet: c.sheet, cell: c.cell, text: c.text ?? '' })))
      if (!got.ok) return refusalOf(got)
      const stop = got.validations?.find((v) => v.severity === 'stop')
      return stop ? await refuseStop(s, stop) : await settle(s, got, boxes)
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

  /** Control+Return: the text read once in the active cell and entered into every cell of the range, which stays selected. */
  async function fill(sheet: string, range: string, text: string, active: string): Promise<EditOutcome> {
    const s = await open()
    const rect = parseRange(range)
    if (!rect) return failed(`Not a range: ${range}`)
    try {
      const got = await s.fillCellText(sheet, range, text, { active })
      if (!got.ok) return refusalOf(got)
      if (got.validation?.severity === 'stop') return await refuseStop(s, { sheet, cell: active, message: got.validation.message })
      const out = await settle(s, got, [{ sheet, rect }])
      return out.ok ? { ...out, select: { sheet, range: rangeAddress(rect) } } : out
    } catch (e) {
      return outcomeOf(e, 'refused')
    }
  }

  /** The session's undo or redo, whole; the selection goes over the cells it took back, on the first of their sheets. */
  async function step(kind: 'undo' | 'redo'): Promise<EditOutcome> {
    const s = await open()
    if (kind === 'redo' && refusedRedo && s.history.redo <= 1) return { ok: true, changes: [] }
    try {
      const got = await s[kind]()
      if (!got) return { ok: true, changes: [] }
      values.clear()
      lists.clear()
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
      let got: EditResult
      if ('columns' in change) {
        const columns = linesOf(change.columns, 'columns')
        if (!columns) return failed(`Not columns: ${change.columns}`)
        got = await s.setLayout([{ sheet: change.sheet, columns, set: { width: change.width, ...(change.width > 0 ? { hidden: null } : {}) } }])
      } else {
        const rows = linesOf(change.rows, 'rows')
        if (!rows) return failed(`Not rows: ${change.rows}`)
        got = await s.setLayout([{ sheet: change.sheet, rows, set: { height: change.height, ...(change.height > 0 ? { hidden: null } : {}) } }])
      }
      // the component already shows the size it set: the cells it changed are enough
      const changes = take(got.changes)
      if (changes.length) emit({ kind: 'changed', changes })
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
      if (needsReload(got)) reload()
      else if (changes.length) emit({ kind: 'changed', changes })
      return { ok: true, changes }
    } catch (e) {
      return outcomeOf(e, 'refused')
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
    edit: change(edit),
    paste: change(paste),
    fill: change(fill),
    undo: serial(() => step('undo')),
    redo: serial(() => step('redo')),
    structure: change(structure),
    layout: change(setLayout),
    format: change(format),
    freeze: change(freeze),
    get history() {
      const h = session?.history ?? { undo: 0, redo: 0 }
      return refusedRedo ? { undo: h.undo, redo: h.redo - 1 } : h
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
    async find(sheet, query, filter, signal) {
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
