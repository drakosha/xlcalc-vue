// What SheetView exposes to its host (undo, redo, format, structure) and its keys for bold, italic and underline: each goes
// through the source as Ctrl+Z does, the selection set without a flash or a Back history entry, the sheet read again after
// a change the window cannot carry, and the sheet displayed afterwards chosen as Excel chooses it.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { toRaw } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { SheetViewApi } from '../src'
import type { EditOutcome, SheetInfo, SheetOp, SheetSource, SheetWindow, SourceEvent, WindowRequest } from '../src/contract/source'
import { active, address, fakeSource, goTo, gridWindow, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const OK: EditOutcome = { ok: true, changes: [] }
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const tab = (w: VueWrapper) => w.find('.xs-tab-on').text()
const backOff = (w: VueWrapper) => w.find('.xs-bar .xs-tool').attributes('disabled') !== undefined

/** Two sheets of a fake source with the given methods; `subscribe` omitted when `quiet` (a source that never reports changes). */
function source(methods: Partial<SheetSource>, quiet = false, sheets: SheetInfo[] = [{ name: 'S', maxRow: 20, maxCol: 6 }, { name: 'T', maxRow: 5, maxCol: 3 }]) {
  const f = fakeSource(sheets)
  Object.assign(f.source, methods)
  if (quiet) delete (f.source as { subscribe?: unknown }).subscribe
  return f
}

describe('SheetView exposes undo and redo', () => {
  it('through the source, with the selection where the outcome puts it: no flash, no Back history entry, on another sheet too', async () => {
    const undo = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [], select: { sheet: 'S', range: 'B2:C3' } }))
    const redo = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [], select: { sheet: 'T', range: 'C4' } }))
    const w = await mountSheet(source({ undo, redo }).source)
    expect(await api(w).undo()).toMatchObject({ ok: true })
    await settle()
    expect([active(w), address(w)]).toEqual(['2:2', 'B2:C3'])
    expect(w.findAll('.xs-flash').length).toBe(0)
    expect(backOff(w)).toBe(true)
    await api(w).redo()
    await settle()
    expect([tab(w), active(w)]).toEqual(['T', '4:3'])
    expect(w.findAll('.xs-flash').length).toBe(0)
    expect(backOff(w)).toBe(true)
  })

  it("return null without the source's methods, and on a read-only sheet", async () => {
    const plain = await mountSheet(source({}).source)
    expect(await api(plain).undo()).toBeNull()
    const undo = vi.fn(async () => OK)
    const ro = await mountSheet(source({ undo }).source, { readOnly: true })
    expect(await api(ro).undo()).toBeNull()
    expect(undo).not.toHaveBeenCalled()
  })
})

describe('a source without edit', () => {
  it('is read-only: no format, structure, freeze, outline or layout from the keys, the menu or the host', async () => {
    const format = vi.fn(async () => OK)
    const structure = vi.fn(async () => OK)
    const freeze = vi.fn(async () => OK)
    const outline = vi.fn(async () => OK)
    const { source: s } = source({ format, structure, freeze, outline })
    delete (s as { edit?: unknown }).edit
    const w = await mountSheet(s)
    await goTo(w, 'B2')
    await key(w, 'b', { ctrlKey: true })
    expect(await api(w).format([{ style: { font: { bold: true } } }])).toBeNull()
    expect(await api(w).structure({ op: 'insertRows', sheet: 'S', row: 2, count: 1 })).toBeNull()
    expect(await api(w).freezePanes()).toBeNull()
    await w.find('.xs-num').trigger('contextmenu', { button: 2, clientX: 20, clientY: 40 })
    await settle()
    expect(document.querySelector('[data-xs-menu="insert"], [data-xs-menu="freeze"]')).toBeNull()
    expect([format, structure, freeze, outline].map((f) => f.mock.calls.length)).toEqual([0, 0, 0, 0])
  })
})

describe('SheetView exposes format', () => {
  it('formats the selection as Excel writes it (entire columns B:B), or the given ranges, on the active sheet', async () => {
    const format = vi.fn(async () => OK)
    const w = await mountSheet(source({ format }).source)
    await w.findAll('.xs-head')[1]!.trigger('click')
    await api(w).format([{ style: { font: { bold: true } } }])
    await api(w).format([{ range: 'A1', numberFormat: 9 }, { style: { fill: null } }])
    expect(format.mock.calls).toEqual([
      ['S', [{ range: 'B:B', style: { font: { bold: true } } }]],
      ['S', [{ range: 'A1', numberFormat: 9 }, { range: 'B:B', style: { fill: null } }]],
    ])
  })

  it('reloads the sheet after a source that reports nothing, and only once after one that sends reload', async () => {
    const quiet = source({ format: vi.fn(async () => OK) }, true)
    const w = await mountSheet(quiet.source)
    const before = quiet.sheetsCalls
    await api(w).format([{ style: { font: { italic: true } } }])
    await settle()
    expect(quiet.sheetsCalls).toBe(before + 1)

    const told = source({})
    told.source.format = vi.fn(async () => (told.emit({ kind: 'reload' }), OK))
    const t = await mountSheet(told.source)
    const was = told.sheetsCalls
    await api(t).format([{ style: { font: { italic: true } } }])
    await settle()
    expect(told.sheetsCalls).toBe(was + 1)
  })

  it('a reload keeps the windows and sizes on screen until the new ones arrive, rather than a default grid in between', async () => {
    let hold = false
    const held: (() => void)[] = []
    const win = (req: WindowRequest): SheetWindow =>
      gridWindow(req, (r, c) => (r === 1 && c === 1 ? 'Region' : ''), { widths: new Array(req.cols).fill(null).map((_, i) => (req.col + i === 1 ? 30 : null)), rowHeights: { 1: 30 } })
    const listeners = new Set<(e: SourceEvent) => void>()
    const src: SheetSource = {
      sheets: async () => [{ name: 'S', maxRow: 20, maxCol: 5 }],
      window: (req) => (hold ? new Promise((done) => held.push(() => done(win(req)))) : Promise.resolve(win(req))),
      subscribe: (l) => (listeners.add(l), () => listeners.delete(l)),
    }
    const w = await mountSheet(src)
    const shown = () => [w.find('[data-cell="1:1"]').text(), (w.findAll('.xs-table colgroup col')[1]!.element as HTMLElement).style.width, (w.find('[data-cell="1:1"]').element.parentElement as HTMLElement).style.height]
    const before = shown()
    expect(before[0]).toBe('Region')
    hold = true
    for (const l of listeners) l({ kind: 'reload' })
    await settle()
    expect(held.length).toBeGreaterThan(0)
    expect(shown()).toEqual(before)
    hold = false
    for (const done of held.splice(0)) done()
    await settle()
    expect(shown()).toEqual(before)
  })

  it('a reload keeps the scroll where it is when rows above the viewport have heights of their own', async () => {
    let hold = false
    const held: (() => void)[] = []
    const tall = Object.fromEntries(Array.from({ length: 40 }, (_, i) => [i + 1, 30]))
    const listeners = new Set<(e: SourceEvent) => void>()
    const src: SheetSource = {
      sheets: async () => [{ name: 'S', maxRow: 400, maxCol: 5 }],
      window: (req) => {
        const win = gridWindow(req, (r, c) => cellAddress(r, c), { rowHeights: Object.fromEntries(Object.entries(tall).filter(([r]) => +r >= req.row && +r < req.row + req.rows)) })
        return hold ? new Promise((done) => held.push(() => done(win))) : Promise.resolve(win)
      },
      subscribe: (l) => (listeners.add(l), () => listeners.delete(l)),
    }
    const w = await mountSheet(src)
    const scroller = w.find('.xs-scroller')
    ;(scroller.element as HTMLElement).scrollTop = 1500
    await scroller.trigger('scroll')
    await settle()
    const top = (scroller.element as HTMLElement).scrollTop
    hold = true
    for (const l of listeners) l({ kind: 'reload' })
    await settle()
    hold = false
    for (const done of held.splice(0)) done()
    await settle()
    expect((scroller.element as HTMLElement).scrollTop).toBe(top)
  })

  it('a format change that the source repaints is not read a second time: only the windows over its ranges are', async () => {
    const painted = source({})
    painted.source.format = vi.fn(async () => (painted.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'B2' }] }), OK))
    const w = await mountSheet(painted.source)
    await goTo(w, 'B2')
    const sheets = painted.sheetsCalls
    const windows = painted.requests.length
    await api(w).format([{ style: { font: { italic: true } } }])
    await settle()
    expect(painted.sheetsCalls).toBe(sheets)
    expect(painted.requests.length - windows).toBe(1)
  })

  it("shows the source's rejection in its message line", async () => {
    const w = await mountSheet(source({ format: vi.fn(async (): Promise<EditOutcome> => ({ ok: false, reason: 'refused', message: 'Лист защищён' })) }).source)
    expect(await api(w).format([{ style: { font: { bold: true } } }])).toMatchObject({ ok: false })
    await settle()
    expect(w.find('.xs-message').text()).toBe('Лист защищён')
  })
})

describe('Ctrl+B, Ctrl+I and Ctrl+U', () => {
  /** A1 bold, B1 plain: the toggle follows the active cell's formatting. */
  function styled() {
    const f = source({ format: vi.fn(async () => OK) })
    const window = f.source.window
    f.source.window = async (req) => {
      const got = await window(req)
      return 'pending' in got ? got : { ...got, styles: [{ font: { bold: true } }], styleIds: { A1: 0 } }
    }
    return f
  }

  it("toggle bold, italic and underline over the selection according to the active cell, Cmd on a Mac, by the key's position on another layout", async () => {
    const f = styled()
    const w = await mountSheet(f.source)
    await key(w, 'b', { ctrlKey: true, code: 'KeyB' })
    await goTo(w, 'B1:C2')
    await key(w, 'и', { ctrlKey: true, code: 'KeyB' })
    await key(w, 'i', { metaKey: true, code: 'KeyI' })
    await key(w, 'г', { ctrlKey: true, code: 'KeyU' })
    expect((f.source.format as ReturnType<typeof vi.fn>).mock.calls).toEqual([
      ['S', [{ range: 'A1', style: { font: { bold: null } } }]],
      ['S', [{ range: 'B1:C2', style: { font: { bold: true } } }]],
      ['S', [{ range: 'B1:C2', style: { font: { italic: true } } }]],
      ['S', [{ range: 'B1:C2', style: { font: { underline: 'single' } } }]],
    ])
  })

  it('do nothing without format, and leave the keys to the browser then', async () => {
    const w = await mountSheet(source({}).source)
    const e = new KeyboardEvent('keydown', { key: 'b', code: 'KeyB', ctrlKey: true, bubbles: true, cancelable: true })
    w.find('.xs-scroller').element.dispatchEvent(e)
    expect(e.defaultPrevented).toBe(false)
  })
})

describe('SheetView exposes structure', () => {
  /** A source whose sheets a structure operation changes, as a workbook would. */
  function book(names: string[], quiet = true) {
    let sheets: SheetInfo[] = names.map((name) => ({ name, maxRow: 5, maxCol: 3 }))
    const ops: SheetOp[] = []
    const f = fakeSource([])
    f.source.sheets = async () => sheets
    f.source.window = async (req) => ({ row: req.row, col: req.col, text: [[req.sheet]], values: [[req.sheet]], formulas: {}, colors: {}, styleIds: {}, styles: [], widths: [null], rowHeights: {}, merges: [] })
    f.source.structure = vi.fn(async (op: SheetOp): Promise<EditOutcome> => {
      ops.push(op)
      const at = (n: string) => sheets.findIndex((x) => x.name === n)
      if (op.op === 'insertSheet') sheets = [...sheets, { name: op.name, maxRow: 0, maxCol: 0 }]
      else if (op.op === 'deleteSheet') sheets = sheets.filter((x) => x.name !== op.sheet)
      else if (op.op === 'hideSheet') sheets = sheets.map((x) => (x.name === op.sheet ? { ...x, hidden: true } : x))
      else if (op.op === 'unhideSheet') sheets = sheets.map((x) => (x.name === op.sheet ? { ...x, hidden: false } : x))
      else if (op.op === 'renameSheet') sheets = sheets.map((x, i) => (i === at(op.sheet) ? { ...x, name: op.name } : x))
      else return { ok: false, reason: 'refused', message: 'Нельзя' }
      if (!quiet) f.emit({ kind: 'reload' })
      return OK
    })
    if (quiet) delete (f.source as { subscribe?: unknown }).subscribe
    return { source: f.source, ops }
  }

  it('displays the sheet Excel displays afterwards: a new or unhidden one, the next after a deleted or hidden one, the previous one at the end', async () => {
    const b = book(['A', 'B', 'C', 'D'])
    const states: { sheet: string }[] = []
    const w = await mountSheet(b.source, { onState: (s: { sheet: string } | null) => s && states.push(s) })
    const shown = async (op: SheetOp) => {
      await api(w).structure(op)
      await settle()
      return states.at(-1)!.sheet
    }
    await w.findAll('.xs-tab')[1]!.trigger('click')
    expect(await shown({ op: 'deleteSheet', sheet: 'B' })).toBe('C')
    expect(await shown({ op: 'hideSheet', sheet: 'C' })).toBe('D')
    expect(await shown({ op: 'hideSheet', sheet: 'D' })).toBe('A')
    expect(await shown({ op: 'insertSheet', name: 'E' })).toBe('E')
    expect(await shown({ op: 'unhideSheet', sheet: 'C' })).toBe('C')
  })

  it("keeps the selection on a renamed sheet, and the source's reload is the only one", async () => {
    const b = book(['A', 'B'], false)
    const w = await mountSheet(b.source)
    await goTo(w, 'B2')
    await api(w).structure({ op: 'renameSheet', sheet: 'A', name: 'Итоги' })
    await settle()
    expect([tab(w), active(w)]).toEqual(['Итоги', '2:2'])
  })

  it('finds the requested sheet by name, case-insensitively, as Excel compares sheet names', async () => {
    const b = book(['A', 'B'])
    const rename = b.source.structure!
    b.source.structure = (op) => rename(op.op === 'renameSheet' ? { ...op, name: op.name.toUpperCase() } : op)
    const w = await mountSheet(b.source)
    await goTo(w, 'B2')
    await api(w).structure({ op: 'renameSheet', sheet: 'A', name: 'Data' })
    await settle()
    expect([tab(w), active(w)]).toEqual(['DATA', '2:2'])
  })

  it('displays a rejection with its message, and does nothing on a read-only sheet', async () => {
    const b = book(['A', 'B'])
    const w = await mountSheet(b.source)
    expect(await api(w).structure({ op: 'moveSheet', sheet: 'A' })).toMatchObject({ ok: false })
    await settle()
    expect(w.find('.xs-message').text()).toBe('Нельзя')
    const ro = await mountSheet(book(['A']).source, { readOnly: true })
    expect(await api(ro).structure({ op: 'insertSheet', name: 'X' })).toBeNull()
  })
})

describe("a source's methods", () => {
  it("are called on the source itself, since a class's methods need their this", async () => {
    const f = fakeSource([{ name: 'S', maxRow: 20, maxCol: 6 }])
    const seen: string[] = []
    const own = { ...f.source }
    const methods = ['undo', 'redo', 'format', 'structure', 'outline', 'freeze', 'explain', 'stats'] as const
    for (const name of methods) {
      ;(own as Record<string, unknown>)[name] = function (this: unknown) {
        seen.push(toRaw(this) === own ? name : `${name}: wrong this`)
        return Promise.resolve(name === 'explain' ? 'A1 = 1' : name === 'stats' ? { cells: 1, numbers: 0, sum: 0, min: null, max: null } : OK)
      }
    }
    const w = await mountSheet(own)
    await api(w).undo()
    await api(w).redo()
    await api(w).format([{ style: { font: { bold: true } } }])
    await api(w).structure({ op: 'insertRows', sheet: 'S', row: 2, count: 1 })
    await api(w).outline({ op: 'group', sheet: 'S', rows: '2:3' })
    await goTo(w, 'B2')
    await api(w).freezePanes()
    api(w).explain()
    await goTo(w, 'A1:B2')
    await new Promise((r) => setTimeout(r, 200))
    await settle()
    expect(new Set(seen)).toEqual(new Set(methods))
  })
})
