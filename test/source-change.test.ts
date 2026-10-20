// A new source is a new workbook, and an unmounted sheet has none: a late answer of the previous source (its sheets, an
// edit's outcome, a rejection) changes nothing on screen and sends nothing to the new source, even from a source that
// ignores the abort signal. Within one workbook the newest list of sheets wins.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, expect, it, vi } from 'vitest'
import { memorySource, SheetView, type MemoryBook, type SheetViewApi } from '../src'
import type { EditOutcome, SheetInfo, SheetOp, SheetSource, SheetWindow, WindowRequest } from '../src/contract/source'
import { cellText, editorKey, fakeSource, goTo, gridWindow, key, mountSheet, settle } from './helpers'

function deferred<T>() {
  let resolve!: (value: T) => void
  let reject!: (e: unknown) => void
  const promise = new Promise<T>((res, rej) => {
    resolve = res
    reject = rej
  })
  return { promise, resolve, reject }
}
const S: SheetInfo = { name: 'S', maxRow: 4, maxCol: 4 }
const tabs = (w: VueWrapper) => w.findAll('.xs-tab').map((x) => x.text())
const alertOpen = () => !!document.querySelector('[data-xs-alert]')
const button = (b: string) => document.querySelector<HTMLElement>(`[data-xs-alert-button="${b}"]`)

let wrappers: VueWrapper[] = []
async function mounted(...args: Parameters<typeof mountSheet>): Promise<VueWrapper> {
  const w = await mountSheet(...args)
  wrappers.push(w)
  return w
}
afterEach(() => {
  for (const w of wrappers) w.unmount()
  wrappers = []
  document.body.innerHTML = ''
})

/** Types `text` into A1 of the old workbook and sends it, then replaces the source while the edit waits. */
async function editThenReplace(w: VueWrapper, text: string, next: ReturnType<typeof fakeSource>): Promise<void> {
  await key(w, text[0]!)
  await w.find('.xs-editor').setValue(text)
  await editorKey(w, 'Enter')
  await settle()
  await w.setProps({ source: next.source })
  await settle()
}

it('a reload of the replaced source does not bring back its sheets', async () => {
  const old = fakeSource([{ name: 'Old', maxRow: 4, maxCol: 4 }])
  const next = fakeSource([{ name: 'New', maxRow: 4, maxCol: 4 }])
  const w = await mounted(old.source)
  const held = deferred<SheetInfo[]>()
  old.source.sheets = () => held.promise
  old.emit({ kind: 'reload' })
  await settle()
  await w.setProps({ source: next.source })
  await settle()
  expect(cellText(w, 1, 1)).toBe('New:A1')
  held.resolve([{ name: 'Old', maxRow: 4, maxCol: 4 }])
  await settle()
  expect(tabs(w)).toEqual(['New'])
  expect(cellText(w, 1, 1)).toBe('New:A1')
})

it("a reload of the replaced source that fails does not show the new workbook's failure", async () => {
  const old = fakeSource([S])
  const next = fakeSource([S], { cell: () => 'new workbook' })
  const w = await mounted(old.source)
  const held = deferred<SheetInfo[]>()
  old.source.sheets = () => held.promise
  old.emit({ kind: 'reload' })
  await settle()
  await w.setProps({ source: next.source })
  await settle()
  held.reject(new Error('the old workbook is gone'))
  await settle()
  expect(w.find('.xs-fallback').exists()).toBe(false)
  expect(cellText(w, 1, 1)).toBe('new workbook')
})

it('keeps the newest list of sheets when two reloads are answered in reverse order', async () => {
  const src = fakeSource([S])
  const w = await mounted(src.source)
  const first = deferred<SheetInfo[]>()
  const second = deferred<SheetInfo[]>()
  let calls = 0
  src.source.sheets = () => (++calls === 1 ? first.promise : second.promise)
  src.emit({ kind: 'reload' })
  src.emit({ kind: 'reload' })
  second.resolve([S, { name: 'New sheet', maxRow: 4, maxCol: 4 }])
  await settle()
  expect(tabs(w)).toEqual(['S', 'New sheet'])
  first.resolve([S])
  await settle()
  expect(tabs(w)).toEqual(['S', 'New sheet'])
})

it('requests no window after it is unmounted while the sheets are read', async () => {
  const held = deferred<SheetInfo[]>()
  const src = fakeSource([S])
  src.source.sheets = () => held.promise
  const w = mount(SheetView, { props: { source: src.source } })
  await settle()
  w.unmount()
  held.resolve([S])
  await settle()
  expect(src.requests).toHaveLength(0)
})

it("an edit of the replaced source does not write its outcome into the new workbook's cells", async () => {
  const held = deferred<EditOutcome>()
  const old = fakeSource([S], { edit: () => held.promise })
  const next = fakeSource([S], { cell: () => 'new workbook' })
  const w = await mounted(old.source)
  await editThenReplace(w, 'x', next)
  expect(old.edits).toHaveLength(1)
  held.resolve({ ok: true, changes: [{ sheet: 'S', cell: 'A1', text: 'old workbook edit', value: 'old workbook edit' }] })
  await settle()
  expect(cellText(w, 1, 1)).toBe('new workbook')
})

it("an edit of the replaced source leaves no pending mark on the new workbook's cell", async () => {
  const held = deferred<EditOutcome>()
  const old = fakeSource([S], { edit: () => held.promise })
  const next = fakeSource([S], { cell: () => 'new workbook' })
  const w = await mounted(old.source)
  await editThenReplace(w, 'x', next)
  expect(w.find('.xs-pending').exists()).toBe(false)
  expect(cellText(w, 1, 1)).toBe('new workbook')
  held.resolve({ ok: true, changes: [] })
  await settle()
})

it("the replaced source's correction of a formula answered late is neither offered nor sent to the new source", async () => {
  const held = deferred<EditOutcome>()
  const old = fakeSource([S], { edit: () => held.promise })
  const next = fakeSource([S], { cell: () => 'new workbook' })
  const w = await mounted(old.source)
  await editThenReplace(w, '=(1+2', next)
  held.resolve({ ok: false, reason: 'invalid', message: 'missing parenthesis', correction: '=(1+2)' })
  await settle()
  expect(alertOpen()).toBe(false)
  button('accept')?.click()
  await settle()
  expect(next.edits).toEqual([])
})

it('closes the correction prompt of the previous workbook when the source is replaced, and its Yes sends nothing', async () => {
  const old = fakeSource([S], { edit: async () => ({ ok: false, reason: 'invalid', message: 'missing parenthesis', correction: '=(1+2)' }) })
  const next = fakeSource([S], { cell: () => 'new workbook' })
  const w = await mounted(old.source)
  await key(w, '=')
  await w.find('.xs-editor').setValue('=(1+2')
  await editorKey(w, 'Enter')
  await settle()
  expect(alertOpen()).toBe(true)
  const yes = button('accept')
  await w.setProps({ source: next.source })
  await settle()
  expect(alertOpen()).toBe(false)
  yes?.click()
  await settle()
  expect(old.edits).toHaveLength(1)
  expect(next.edits).toEqual([])
})

/** An in-memory workbook whose `structure` is a mock; `hold` keeps the window requests it matches until the test answers them. */
function withStructure(sheets: MemoryBook['sheets'], hold: (req: WindowRequest) => boolean = () => false) {
  const source = memorySource({ sheets }) as SheetSource
  const structure = vi.fn(async (_op: SheetOp): Promise<EditOutcome> => ({ ok: true, changes: [] }))
  source.structure = structure
  const read = source.window.bind(source)
  const held: { req: WindowRequest; answer: (w: SheetWindow) => void }[] = []
  source.window = (req, signal) => (hold(req) ? new Promise((answer) => held.push({ req, answer })) : read(req, signal))
  return { source, structure, held }
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi

it("a merge whose cells the replaced source was still reading is not sent to the new source", async () => {
  const old = withStructure([{ name: 'S', cells: { A3000: 1 } }], (req) => req.rows > 1000)
  const next = withStructure([{ name: 'S', cells: { A3000: 1 } }])
  const w = await mounted(old.source)
  await goTo(w, 'A1:A3000')
  const merged = api(w).merge('mergeCells')
  await settle()
  expect(old.held).toHaveLength(1)
  await w.setProps({ source: next.source })
  await settle()
  old.held[0]!.answer(gridWindow(old.held[0]!.req, () => ''))
  await settle()
  button('accept')?.click()
  await expect(merged).resolves.toBeNull()
  expect(next.structure).not.toHaveBeenCalled()
  expect(old.structure).not.toHaveBeenCalled()
})

it("closes the merge warning of the previous workbook when the source is replaced, and its OK sends nothing", async () => {
  const old = withStructure([{ name: 'S', cells: { A1: 'a', B1: 'b' } }])
  const next = withStructure([{ name: 'S', cells: { A1: 'a', B1: 'b' } }])
  const w = await mounted(old.source)
  await goTo(w, 'A1:B1')
  const merged = api(w).merge('mergeCells')
  await settle()
  expect(alertOpen()).toBe(true)
  const ok = button('accept')
  await w.setProps({ source: next.source })
  await settle()
  expect(alertOpen()).toBe(false)
  ok?.click()
  await expect(merged).resolves.toBeNull()
  expect(next.structure).not.toHaveBeenCalled()
  expect(old.structure).not.toHaveBeenCalled()
})

it("a double click on the fill handle whose column the replaced source was still reading fills nothing in the new workbook", async () => {
  const cells = { A1: 1, A2: 2, A3: 3, A4: 4 }
  const old = withStructure([{ name: 'S', cells }], (req) => req.cols === 1 && req.col === 1 && req.row === 2)
  const next = withStructure([{ name: 'S', cells }])
  const w = await mounted(old.source)
  await goTo(w, 'B1')
  await w.find('[data-xs-fill-handle]').trigger('dblclick')
  await settle()
  expect(old.held).toHaveLength(1)
  await w.setProps({ source: next.source })
  await settle()
  old.held[0]!.answer(gridWindow(old.held[0]!.req, (r) => String(r)))
  await settle()
  expect(next.structure).not.toHaveBeenCalled()
  expect(old.structure).not.toHaveBeenCalled()
})

it("a structure operation of the replaced source answered late does not move the new workbook to its sheet", async () => {
  let answer!: (o: EditOutcome) => void
  const old = withStructure([{ name: 'S', cells: {} }])
  old.structure.mockImplementation(() => new Promise((r) => (answer = r)))
  const next = withStructure([{ name: 'S', cells: {} }, { name: 'Added', cells: {} }])
  const w = await mounted(old.source)
  const inserted = api(w).structure({ op: 'insertSheet', name: 'Added' })
  await settle()
  await w.setProps({ source: next.source })
  await settle()
  expect(w.find('.xs-tab-on').text()).toBe('S')
  answer({ ok: true, changes: [] })
  await inserted
  await settle()
  expect(w.find('.xs-tab-on').text()).toBe('S')
})
