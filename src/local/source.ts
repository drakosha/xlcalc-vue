// localSource: a sheet source over an xlcalc book whose session runs in a Web Worker pool or in the page's thread.
// The session holds the book and computes; this side keeps the layout a window needs and the sheets' shown values for
// search and the display filter, and patches both as it edits.
import { createInlineCalc, LOCALES, type AsyncSession, type BookInput, type Calc, type CellChange, type CellReplace, type CellView, type LayoutEdit, type PatchFields } from 'xlcalc'
import { cellAddress, parseCell, parseRect } from '../core/a1'
import type { CellEntry, CellUpdate, EditOutcome, LayoutChange, SheetFilter, SheetInfo, SheetOp, SourceEvent, WindowRequest } from '../source'
import { parseTarget, typedInput } from './input'
import { grow, lineLook, readLayout, sheetOf, type BookLayout, type SheetLayout } from './layout'
import { columnValuesOf, filterActive, filterKey, findIn, SheetValues, statsOf, visibleList } from './query'
import type { LocalSource, LocalSourceOptions } from './types'
import { loadLocale } from './locale'
import { Formulas, windowOf } from './window'

const MAX_CELLS = 20_000

const abortError = () => new DOMException('The request was aborted', 'AbortError')
function check(signal?: AbortSignal): void {
  if (signal?.aborted) throw abortError()
}
const failed = (message: string): EditOutcome => ({ ok: false, reason: 'failed', message })
const sheetMissing = (name: string): EditOutcome => failed(`No sheet ${name}`)

/** An exception of the session as an outcome: input the book refuses, a busy calculator, anything else failed. */
function outcomeOf(e: unknown, input: 'refused' | 'failed'): EditOutcome {
  const err = e as { name?: string; code?: string; message?: string }
  const message = err?.message ?? String(e)
  if (err?.name === 'InputError') return { ok: false, reason: input, message }
  if (err?.name === 'CalcError' && (err.code === 'soft_timeout' || err.code === 'queue_full' || err.code === 'aborted')) return { ok: false, reason: 'busy', message }
  return failed(message)
}

async function startCalc(opts: LocalSourceOptions): Promise<Calc> {
  const kind = typeof opts.calc === 'string' ? opts.calc : typeof Worker === 'function' ? 'pool' : 'inline'
  if (kind === 'inline') return createInlineCalc()
  const { createCalcPool } = await import('xlcalc/worker')
  return createCalcPool({ threads: 1, ...opts.pool })
}

/** A source over `input`: the book goes to the session as it is (its cells are not copied), so do not change it afterwards. */
export function localSource(input: BookInput, opts: LocalSourceOptions = {}): LocalSource {
  let layout: BookLayout = readLayout(input)
  // Formulas show and text sorts in the book's locale once it is loaded; an unknown locale fails the opening.
  let formulas: Formulas
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
    formulas = new Formulas(LOCALES.get(layout.locale)!)
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
      const key = `${ch.sheet}!${ch.cell}`
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

  /** Cells by address over the rectangles that hold `cells`, one fetch per sheet. */
  async function viewsOf(s: AsyncSession, cells: { sheet: string; row: number; col: number }[]): Promise<Map<string, Record<string, CellView>>> {
    const boxes = new Map<string, { r1: number; c1: number; r2: number; c2: number }>()
    for (const { sheet, row, col } of cells) {
      const b = boxes.get(sheet)
      boxes.set(sheet, b ? { r1: Math.min(b.r1, row), c1: Math.min(b.c1, col), r2: Math.max(b.r2, row), c2: Math.max(b.c2, col) } : { r1: row, c1: col, r2: row, c2: col })
    }
    const out = new Map<string, Record<string, CellView>>()
    await Promise.all([...boxes].map(async ([sheet, b]) => out.set(sheet, await s.getCells(sheet, `${cellAddress(b.r1, b.c1)}:${cellAddress(b.r2, b.c2)}`))))
    return out
  }

  async function edit(entries: CellEntry[]): Promise<EditOutcome> {
    const s = await open()
    const cells: { sheet: string; row: number; col: number; l: SheetLayout; addr: string; text: string | null }[] = []
    for (const e of entries) {
      const l = sheetOf(layout, e.sheet)
      const at = parseCell(e.cell)
      if (!l || !at) return failed(`No cell ${e.sheet}!${e.cell}`)
      cells.push({ sheet: l.name, row: at.row, col: at.col, l, addr: cellAddress(at.row, at.col), text: e.text })
    }
    if (!cells.length) return { ok: true, changes: [] }
    try {
      const before = await viewsOf(s, cells)
      const now = s.clock.now
      const edits: CellReplace[] = []
      for (const c of cells) {
        const got = typedInput(c.text, before.get(c.sheet)?.[c.addr], { locale: layout.locale, dateSystem: layout.dateSystem, now, inherited: lineLook(c.l, c.row, c.col).numberFormat })
        if (!got.ok) return { ok: false, reason: 'invalid', message: got.message, at: got.at, cells: [{ sheet: c.sheet, cell: c.addr }] }
        edits.push({ sheet: c.sheet, cell: c.addr, input: got.input })
      }
      const { changes } = await s.replaceCells(edits)
      // The edited cells as the session now shows them, with their formulas: a value that did not change is no change.
      const after = await viewsOf(s, cells)
      const edited = new Map<string, CellUpdate>()
      for (const c of cells) {
        const v = after.get(c.sheet)?.[c.addr]
        edited.set(`${c.sheet}!${c.addr}`, { sheet: c.sheet, cell: c.addr, text: v?.text ?? '', value: v ? v.value : null, ...(v?.color ? { color: v.color } : {}), formula: formulas.of(c.l, c.row, c.col, v) })
      }
      const updates = take(changes, edited)
      emit({ kind: 'changed', changes: updates })
      // A header cell names its table's column: formulas that use it are written anew.
      if (cells.some((c) => c.l.tables.some((t) => t.headerRow === c.row && c.col >= t.rect.c1 && c.col <= t.rect.c2))) emit({ kind: 'reload' })
      return { ok: true, changes: updates }
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
