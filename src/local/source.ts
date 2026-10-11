// localSource: a sheet source over an xlcalc book whose session runs in a Web Worker pool or in the page's thread.
// The session holds the book and computes; this side keeps the layout a window needs and the sheets' shown values for
// search and the display filter, and patches both as it edits.
import { createInlineCalc, type AsyncSession, type BookInput, type Calc, type CellChange, type EditResult, type LayoutEdit, type PatchFields } from 'xlcalc'
import { cellAddress, parseCell, parseRange, rangeAddress } from 'xlcalc/a1'
import { parseRect } from '../core/a1'
import type { CellEntry, CellUpdate, EditOutcome, LayoutChange, SheetFilter, SheetInfo, SheetOp, SourceEvent, WindowRequest } from '../source'
import { parseTarget } from './input'
import { grow, readLayout, sheetOf, type BookLayout, type SheetLayout } from './layout'
import { columnValuesOf, filterActive, filterKey, findIn, SheetValues, statsOf, visibleList } from './query'
import type { LocalSource, LocalSourceOptions } from './types'
import { loadLocale } from './locale'
import { boxesOf, cellKey, failed, needsReload, outcomeOf, refusalOf, updateOf, type CellsBox } from './outcome'
import { windowOf } from './window'

const MAX_CELLS = 20_000

const abortError = () => new DOMException('The request was aborted', 'AbortError')
function check(signal?: AbortSignal): void {
  if (signal?.aborted) throw abortError()
}
const sheetMissing = (name: string): EditOutcome => failed(`No sheet ${name}`)

async function startCalc(opts: LocalSourceOptions): Promise<Calc> {
  const kind = typeof opts.calc === 'string' ? opts.calc : typeof Worker === 'function' ? 'pool' : 'inline'
  if (kind === 'inline') return createInlineCalc()
  const { createCalcPool } = await import('xlcalc/worker')
  return createCalcPool({ threads: 1, ...opts.pool })
}

/** A source over `input`: the book goes to the session as it is (its cells are not copied), so do not change it afterwards. */
export function localSource(input: BookInput, opts: LocalSourceOptions = {}): LocalSource {
  let layout: BookLayout = readLayout(input)
  // Text sorts in the book's locale once it is loaded; an unknown locale fails the opening.
  let collator: Intl.Collator
  const maxCells = opts.maxCells ?? MAX_CELLS
  const listeners = new Set<(event: SourceEvent) => void>()
  /** Shown values of sheets that search or a filter needed, by sheet name. */
  const values = new Map<string, Promise<SheetValues>>()
  /** The list of rows of a sheet's display filter, as last given by `visibleRows`: windows under the filter read it. */
  const lists = new Map<string, { key: string; rows: number[] }>()
  let calc: Calc | null = null
  let session: AsyncSession | null = null

  /** The source's own changes one after another: each reads cells, then writes them, and must not interleave. */
  let last: Promise<unknown> = Promise.resolve()
  function serial<A extends unknown[], R>(run: (...args: A) => Promise<R>): (...args: A) => Promise<R> {
    return (...args) => {
      const next = last.then(() => run(...args))
      last = next.catch(() => {})
      return next
    }
  }

  const ready = (async () => {
    await loadLocale(layout.locale)
    calc = typeof opts.calc === 'object' ? opts.calc : await startCalc(opts)
    const s = await calc.createSession(input, opts.session)
    session = s
    collator = new Intl.Collator(layout.locale, { numeric: true, sensitivity: 'base' })
    await growBySpills(s, input)
  })()
  ready.catch(() => {})
  const open = async (): Promise<AsyncSession> => {
    await ready
    return session!
  }

  function emit(event: SourceEvent): void {
    for (const l of listeners) l(event)
  }

  /** Arrays may spill past the cells the book holds: the sheet's extent takes in the spill ranges of its formulas. */
  async function growBySpills(s: AsyncSession, book: BookInput): Promise<void> {
    await Promise.all(
      book.sheets.map(async (sheet) => {
        const l = sheetOf(layout, sheet.name)
        let box: { r1: number; c1: number; r2: number; c2: number } | null = null
        for (const [addr, cell] of Object.entries(sheet.cells)) {
          if (cell.formula === undefined) continue
          const at = parseCell(addr)
          if (!at) continue
          box = box ? { r1: Math.min(box.r1, at.row), c1: Math.min(box.c1, at.col), r2: Math.max(box.r2, at.row), c2: Math.max(box.c2, at.col) } : { r1: at.row, c1: at.col, r2: at.row, c2: at.col }
        }
        if (!l || !box) return
        const views = await s.getCells(sheet.name, `${cellAddress(box.r1, box.c1)}:${cellAddress(box.r2, box.c2)}`)
        for (const v of Object.values(views)) {
          const spill = v.spill ? parseRect(v.spill) : null
          if (spill) grow(l, spill.r2, spill.c2)
        }
      }),
    )
  }

  /** The layout again from the session's book, after an operation that moves things about; kept values are dropped. */
  async function reread(s: AsyncSession): Promise<void> {
    const book = await s.getBook()
    layout = readLayout(book)
    values.clear()
    lists.clear()
    await growBySpills(s, book)
  }

  /** The session's changes as cell updates, and the kept values and extents brought up to date with them. */
  function take(changes: CellChange[], edited: Map<string, CellUpdate> = new Map()): CellUpdate[] {
    const out = new Map(edited)
    for (const ch of changes) {
      const key = cellKey(ch.sheet, ch.cell)
      if (!out.has(key)) out.set(key, { sheet: ch.sheet, cell: ch.cell, text: ch.text ?? '', value: ch.value, ...(ch.color ? { color: ch.color } : {}) })
    }
    const list = [...out.values()]
    for (const u of list) {
      const l = sheetOf(layout, u.sheet)
      const at = parseCell(u.cell)
      if (!l || !at) continue
      if (u.text !== '' || (u.value !== null && u.value !== '')) grow(l, at.row, at.col)
      void values.get(l.name)?.then((v) => v.set(at.row, at.col, u.text, u.value))
    }
    return list
  }

  function sheetValues(s: AsyncSession, l: SheetLayout): Promise<SheetValues> {
    let got = values.get(l.name)
    if (!got) {
      const range = l.maxRow && l.maxCol ? `A1:${cellAddress(l.maxRow, l.maxCol)}` : null
      got = range ? s.getCells(l.name, range).then((views) => new SheetValues(views)) : Promise.resolve(new SheetValues({}))
      values.set(l.name, got)
      got.catch(() => values.delete(l.name))
    }
    return got
  }

  /** The rows of a sheet under a filter: the list `visibleRows` gave for it, so windows match what the component holds. */
  async function rowsUnder(s: AsyncSession, l: SheetLayout, filter: SheetFilter, fresh: boolean): Promise<number[]> {
    const key = filterKey(filter)
    const kept = lists.get(l.name)
    if (!fresh && kept?.key === key) return kept.rows
    const rows = visibleList(await sheetValues(s, l), filter, l.maxRow, collator)
    lists.set(l.name, { key, rows })
    return rows
  }

  function info(l: SheetLayout): SheetInfo {
    const f = l.frozen
    return { name: l.name, maxRow: l.maxRow, maxCol: l.maxCol, ...(f ? { freeze: { rows: f.rows ?? 0, cols: f.columns ?? 0 } } : {}) }
  }

  function sheetFor(name: string): SheetLayout {
    const l = sheetOf(layout, name)
    if (!l) throw new Error(`No sheet ${name}`)
    return l
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

  /** The book changed in many places: kept values are dropped and subscribers read the sheets and windows again. */
  async function reload(s: AsyncSession): Promise<void> {
    await reread(s)
    emit({ kind: 'reload' })
  }

  /** A change made: its updates, the edited cells' as they are now, told to subscribers; a reload where updates cannot carry it. */
  async function settle(s: AsyncSession, got: EditResult, boxes: CellsBox[]): Promise<EditOutcome> {
    const { updates, truncated } = await editedIn(s, boxes)
    const changes = take(got.changes, updates)
    if (truncated || needsReload(got)) await reload(s)
    else if (changes.length) emit({ kind: 'changed', changes })
    return { ok: true, changes }
  }

  /** A value a stop rule of data validation does not take: the call is taken back whole, as Excel takes it back. */
  async function refuseStop(s: AsyncSession, alert: { sheet: string; cell: string; message: string }): Promise<EditOutcome> {
    await s.undo()
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

  async function structure(op: SheetOp): Promise<EditOutcome> {
    const s = await open()
    try {
      const got = await s.editStructure(op)
      if (!got.ok) return { ok: false, reason: 'refused', message: got.message }
      await reread(s)
      const changes = take(got.changes)
      emit({ kind: 'reload' })
      return { ok: true, changes }
    } catch (e) {
      return outcomeOf(e, 'failed')
    }
  }

  async function setLayout(change: LayoutChange): Promise<EditOutcome> {
    const s = await open()
    const l = sheetOf(layout, change.sheet)
    if (!l) return sheetMissing(change.sheet)
    const span = 'columns' in change ? change.columns : change.rows
    const target = parseTarget(span.includes(':') ? span : `${span}:${span}`)
    try {
      let edit: LayoutEdit
      if ('columns' in change && target?.kind === 'columns') {
        const set = { width: change.width, ...(change.width > 0 ? { hidden: null } : {}) }
        edit = { sheet: l.name, columns: target.span, set }
        const { changes } = await s.setLayout([edit])
        l.columns.patch(target.lo, target.hi, set)
        return done(take(changes))
      }
      if ('rows' in change && target?.kind === 'rows') {
        const set = { height: change.height, ...(change.height > 0 ? { hidden: null } : {}) }
        edit = { sheet: l.name, rows: target.span, set }
        const { changes } = await s.setLayout([edit])
        l.rows.patch(target.lo, target.hi, set)
        return done(take(changes))
      }
      return failed(`Not columns or rows: ${span}`)
    } catch (e) {
      return outcomeOf(e, 'failed')
    }
  }
  function done(changes: CellUpdate[]): EditOutcome {
    if (changes.length) emit({ kind: 'changed', changes })
    return { ok: true, changes }
  }

  async function format(sheet: string, range: string, patch: Parameters<LocalSource['format']>[2]): Promise<EditOutcome> {
    const s = await open()
    const set: PatchFields = {}
    if (patch.style !== undefined) set.style = patch.style
    if (patch.numberFormat !== undefined) set.numberFormat = patch.numberFormat
    if (!Object.keys(set).length) return { ok: true, changes: [] }
    try {
      const got = await s.patchCells([{ sheet, range, set }])
      if (got.layout || got.styles) await reread(s)
      const updates = take(got.changes)
      if (set.style !== undefined) emit({ kind: 'reload' })
      else if (updates.length) emit({ kind: 'changed', changes: updates })
      return { ok: true, changes: updates }
    } catch (e) {
      return outcomeOf(e, 'refused')
    }
  }

  const source: LocalSource = {
    maxCells,
    locale: layout.locale,
    ready,
    async sheets(signal) {
      await ready
      check(signal)
      return layout.sheets.map(info)
    },
    async window(req: WindowRequest, signal) {
      const s = await open()
      check(signal)
      const l = sheetFor(req.sheet)
      const cols = Math.max(0, req.cols)
      const count = Math.max(0, Math.min(req.rows, Math.floor(maxCells / Math.max(1, cols))))
      const filtered = filterActive(req.filter)
      let rows: number[]
      if (filtered) rows = (await rowsUnder(s, l, req.filter!, false)).slice(req.row - 1, req.row - 1 + count)
      else rows = Array.from({ length: count }, (_, i) => req.row + i)
      // the viewport takes rows in ascending order; the window keeps the display order
      const asked = [...rows].sort((a, b) => a - b)
      const top = asked[0] ?? 1
      const range = `${cellAddress(top, req.col)}:${cellAddress(asked[asked.length - 1] ?? top, req.col + Math.max(1, cols) - 1)}`
      const v = await s.getViewport(l.name, range, { formulaLocal: true, editText: true, maxCells, ...(filtered || !rows.length ? { rows: asked } : {}) })
      check(signal)
      return windowOf(v, { row: req.row, rows, col: req.col, cols, filtered, notice: opts.notice })
    },
    edit: serial(edit),
    structure: serial(structure),
    layout: serial(setLayout),
    format: serial(format),
    freeze: serial(async (sheet: string, panes: { rows: number; cols: number } | null): Promise<EditOutcome> => {
      const s = await open()
      const l = sheetOf(layout, sheet)
      if (!l) return sheetMissing(sheet)
      const frozen = panes && (panes.rows > 0 || panes.cols > 0) ? { ...(panes.rows > 0 ? { rows: panes.rows } : {}), ...(panes.cols > 0 ? { columns: panes.cols } : {}) } : null
      try {
        const { changes } = await s.setLayout([{ sheet: l.name, frozen }])
        if (frozen) l.frozen = frozen
        else delete l.frozen
        emit({ kind: 'reload' })
        return { ok: true, changes: take(changes) }
      } catch (e) {
        return outcomeOf(e, 'failed')
      }
    }),
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    async explain(sheet, cell) {
      const s = await open()
      return (await s.explain(sheetFor(sheet).name, cell)).text
    },
    async stats(sheet, range, filter, signal) {
      const s = await open()
      const l = sheetFor(sheet)
      const target = parseTarget(range)
      if (!target) throw new Error(`Not a range: ${range}`)
      const text = target.kind === 'cells' ? `${cellAddress(target.rect.r1, target.rect.c1)}:${cellAddress(target.rect.r2, target.rect.c2)}` : target.range
      const [views, rows] = await Promise.all([s.getCells(l.name, text), filterActive(filter) ? rowsUnder(s, l, filter, false) : null])
      check(signal)
      return statsOf(views, rows ? new Set(rows) : null)
    },
    async find(sheet, query, filter, signal) {
      const s = await open()
      const l = sheetFor(sheet)
      const [v, rows] = await Promise.all([sheetValues(s, l), filterActive(filter) ? rowsUnder(s, l, filter, false) : null])
      check(signal)
      return findIn(v, query, rows ? new Set(rows) : null)
    },
    async visibleRows(sheet, filter, signal) {
      const s = await open()
      const l = sheetFor(sheet)
      if (!filterActive(filter)) {
        lists.delete(l.name)
        return { rows: null, total: l.maxRow }
      }
      const rows = await rowsUnder(s, l, filter, true)
      check(signal)
      return { rows, total: l.maxRow }
    },
    async columnValues(sheet, col, filter, signal) {
      const s = await open()
      const l = sheetFor(sheet)
      const v = await sheetValues(s, l)
      check(signal)
      return columnValuesOf(v, col, filter, l.maxRow, collator)
    },
    async book() {
      return (await open()).getBook()
    },
    async result() {
      return (await open()).getResult()
    },
    async close() {
      await ready.catch(() => {})
      await session?.close()
      if (calc && typeof opts.calc !== 'object') await calc.close()
    },
  }
  return source
}
