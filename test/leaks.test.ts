// What the sheet leaves behind when it goes or shows another sheet: no search after its pause, requests out aborted,
// late answers dropped, and reads of the fill handle and of the merge's warning stopped.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { sheetContextKey } from '../src/components/context'
import type { SheetContext } from '../src/core/sheet'
import type { EditOutcome, SheetOp, SheetSource, WindowRequest } from '../src/contract/source'
import { fakeSource, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
  document.body.innerHTML = ''
  localStorage.clear()
})

const OK: EditOutcome = { ok: true, changes: [] }
const contextOf = (w: VueWrapper) => (w.vm.$ as unknown as { provides: Record<symbol, SheetContext> }).provides[sheetContextKey as symbol]!
const fakeTimers = () => vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'Date'] })

/** A request held until the test lets it go, with the signal it came with. */
interface Held {
  signal?: AbortSignal
  go: () => void
}
/** A promise held with its signal. */
function hold<T>(list: Held[], signal: AbortSignal | undefined, answer: () => Promise<T>): Promise<T> {
  return new Promise((resolve, reject) => list.push({ signal, go: () => void answer().then(resolve, reject) }))
}
/** The source's windows held from `on()` on. */
function holdWindows(source: SheetSource) {
  const held: (Held & { req: WindowRequest })[] = []
  const own = source.window.bind(source)
  let on = false
  source.window = (req, signal) => (on ? new Promise((resolve, reject) => held.push({ req, signal, go: () => void own(req).then(resolve, reject) })) : own(req, signal))
  return { held, on: () => (on = true) }
}

describe('find', () => {
  const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'apple' } }] })

  it('asks nothing once the sheet is gone during the pause after typing', async () => {
    fakeTimers()
    const source = memorySource(book())
    const find = vi.spyOn(source, 'find')
    const w = await mountSheet(source)
    await key(w, 'f', { ctrlKey: true })
    await settle()
    await w.find('.xs-find-input').setValue('apple')
    w.unmount()
    await vi.advanceTimersByTimeAsync(1000)
    expect(find).not.toHaveBeenCalled()
  })

  it('aborts a search out when the sheet goes', async () => {
    fakeTimers()
    const held: Held[] = []
    const source = memorySource(book())
    const own = source.find!.bind(source)
    source.find = (req, signal) => hold(held, signal, () => own(req))
    const w = await mountSheet(source)
    await key(w, 'f', { ctrlKey: true })
    await settle()
    await w.find('.xs-find-input').setValue('apple')
    await vi.advanceTimersByTimeAsync(300)
    expect(held).toHaveLength(1)
    w.unmount()
    expect(held[0]!.signal?.aborted).toBe(true)
  })
})

describe('the display filter', () => {
  it('aborts the list of rows out when the sheet goes', async () => {
    const held: Held[] = []
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 'h', A2: 'a', A3: 'b' } }] })
    const own = source.visibleRows!.bind(source)
    source.visibleRows = (sheet, filter, signal) => hold(held, signal, () => own(sheet, filter))
    const w = await mountSheet(source)
    contextOf(w).filter.cols.value = { 1: { values: ['a'] } }
    await settle()
    expect(held).toHaveLength(1)
    w.unmount()
    expect(held[0]!.signal?.aborted).toBe(true)
  })
})

describe('explain', () => {
  it('drops an answer that comes after the sheet went', async () => {
    const held: Held[] = []
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: { value: 2, formula: 'B1*2' }, B1: 1 } }] })
    source.explain = () => hold(held, undefined, async () => 'A1 = 2  ← B1*2')
    const w = await mountSheet(source)
    const ctx = contextOf(w)
    ctx.api.explain()
    await settle()
    expect(held).toHaveLength(1)
    w.unmount()
    held[0]!.go()
    await settle()
    expect(ctx.explain.lines.value).toEqual([])
  })
})

describe('reads of the fill handle and the merge warning', () => {
  it('a double click on the handle reads no further and fills nothing once another sheet is shown', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 1, A2: 2, A3: 3, B1: 5 } }, { name: 'T', cells: { A1: 1 } }] })
    const structure = vi.fn(async (_op: SheetOp) => OK)
    source.structure = structure
    const windows = holdWindows(source)
    const w = await mountSheet(source)
    await goTo(w, 'B1')
    windows.on()
    await w.find('[data-xs-fill-handle]').trigger('dblclick')
    await settle()
    expect(windows.held).toHaveLength(1)
    await contextOf(w).pickSheet(1)
    await settle()
    expect(windows.held[0]!.signal?.aborted).toBe(true)
    windows.held[0]!.go()
    await settle()
    expect(structure).not.toHaveBeenCalled()
  })

  it('the merge warning stops reading the windows when the sheet goes', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 1, A4000: 2 } }] })
    source.structure = vi.fn(async () => OK)
    const windows = holdWindows(source)
    const w = await mountSheet(source)
    await goTo(w, 'A1:A5000')
    windows.on()
    void contextOf(w).api.merge('mergeCells')
    await settle()
    expect(windows.held).toHaveLength(1)
    w.unmount()
    expect(windows.held[0]!.signal?.aborted).toBe(true)
  })
})

describe('a long press', () => {
  it('held on a tab while the sheet goes opens nothing after it', async () => {
    fakeTimers()
    const f = fakeSource([{ name: 'S', maxRow: 5, maxCol: 5 }, { name: 'T', maxRow: 5, maxCol: 5 }])
    f.source.structure = vi.fn(async () => OK)
    const w = await mountSheet(f.source)
    await w.findAll('.xs-tab')[1]!.trigger('pointerdown', { pointerType: 'touch', isPrimary: true, button: 0, clientX: 20, clientY: 30 })
    const asked = f.requests.length
    w.unmount()
    await vi.advanceTimersByTimeAsync(1000)
    expect(f.requests.slice(asked).map((r) => r.sheet)).toEqual([])
  })
})
