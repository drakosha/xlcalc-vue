// Copy, cut and paste inside the book, as Excel's: Control+C and Control+X mark the range with a moving border, a paste
// of the same text pastes the range itself through `structure` (copyRange, moveRange), other text is pasted as text; the
// cells' menu (a right click, a finger's long press) cuts, copies and pastes through the browser's clipboard API.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, ruMessages, type MemoryBook } from '../src'
import { localSource } from '../src/local'
import type { EditOutcome, SheetOp, SheetWindow } from '../src/contract/source'
import { address, cellText, editorKey, goTo, key, mountSheet, paste, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
  Reflect.deleteProperty(navigator, 'clipboard')
})

const OK: EditOutcome = { ok: true, changes: [] }
const book = (): MemoryBook => ({
  sheets: [
    { name: 'S', cells: { A1: 1, A2: 2, A3: 3, B1: { value: 10, formula: 'A1*10' }, C1: 'x' }, locked: [{ range: 'F1:F9', reason: 'script' }] },
    { name: 'T', cells: {} },
  ],
})
/** A memory book with a structure method the test watches, and its edits. */
function withStructure(answer: (op: SheetOp) => EditOutcome = () => OK) {
  const source = memorySource(book())
  const structure = vi.fn(async (op: SheetOp) => answer(op))
  source.structure = structure
  const edit = vi.spyOn(source, 'edit')
  return { source, structure, edit }
}
/** Copy or cut on the grid as the browser sends it; the text the sheet put on the clipboard. */
function clipboard(w: VueWrapper, kind: 'copy' | 'cut'): string | undefined {
  const data: Record<string, string> = {}
  const e = new Event(kind, { bubbles: true, cancelable: true }) as ClipboardEvent
  Object.defineProperty(e, 'clipboardData', { value: { setData: (t: string, v: string) => (data[t] = v) } })
  w.find('.xs-scroller').element.dispatchEvent(e)
  return data['text/plain']
}
/** A paste whose clipboard the browser does not give. */
function pasteUnread(w: VueWrapper): void {
  w.find('.xs-scroller').element.dispatchEvent(new Event('paste', { bubbles: true, cancelable: true }))
}
const ants = (w: VueWrapper) => w.find('[data-xs-ants]').exists()
const ops = (structure: ReturnType<typeof withStructure>['structure']) => structure.mock.calls.map((c) => c[0])

describe('the clipboard’s plain text, as Excel copies it (xlcalc’s formatClipboard)', () => {
  it('rows end with CRLF, a field with a tab or a line break goes in quotes, a hidden column is left out', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 'a\tb', B1: 'hidden', C1: 'say "hi"', A2: 'x\ny', C2: 3 }, widths: { B: 0 } }] })
    const w = await mountSheet(source)
    await goTo(w, 'A1:C2')
    expect(clipboard(w, 'copy')).toBe('"a\tb"\tsay "hi"\r\n"x\ny"\t3\r\n')
  })

  it('on a Mac the text has no CRLF at the end, as Excel for Mac copies', async () => {
    vi.spyOn(navigator, 'platform', 'get').mockReturnValue('MacIntel')
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'A1:A2')
    expect(clipboard(w, 'copy')).toBe('1\r\n2')
  })
})

describe('Control+C and Control+V inside the book', () => {
  it('a copy marks the range; a paste of its text copies the range from the top-left cell and selects what it pasted; the mark stays', async () => {
    const { source, structure, edit } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1:B2')
    expect(ants(w)).toBe(false)
    const text = clipboard(w, 'copy')
    expect(text).toBe('1\t10\r\n2\t\r\n')
    await settle()
    expect(ants(w)).toBe(true)
    await goTo(w, 'D5')
    paste(w, text!)
    await settle()
    expect(ops(structure)).toEqual([{ op: 'copyRange', from: { sheet: 'S', range: 'A1:B2' }, to: { sheet: 'S', range: 'D5' } }])
    expect(edit).not.toHaveBeenCalled()
    expect(address(w)).toBe('D5:E6')
    expect(ants(w)).toBe(true)
  })

  it('tiles a selection the copy’s size divides; another selection gets one copy from its top-left cell', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1:B1')
    const text = clipboard(w, 'copy')!
    await goTo(w, 'D1:E3')
    paste(w, text)
    await settle()
    expect(address(w)).toBe('D1:E3')
    await goTo(w, 'D5:F6')
    paste(w, text)
    await settle()
    expect(address(w)).toBe('D5:E5')
    expect(ops(structure).map((op) => (op as { to: unknown }).to)).toEqual([
      { sheet: 'S', range: 'D1:E3' },
      { sheet: 'S', range: 'D5' },
    ])
  })

  it('pastes on another sheet, and a clipboard the browser does not give pastes the copy', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B1')
    clipboard(w, 'copy')
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    await goTo(w, 'C3')
    pasteUnread(w)
    await settle()
    expect(ops(structure)).toEqual([{ op: 'copyRange', from: { sheet: 'S', range: 'B1' }, to: { sheet: 'T', range: 'C3' } }])
    expect(ants(w)).toBe(false)
    await w.findAll('.xs-tab')[0]!.trigger('click')
    await settle()
    expect(ants(w)).toBe(true)
  })

  it('a cut moves the range to the selection’s top-left cell, its mark goes, and its text is not pasted again', async () => {
    const { source, structure, edit } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1:A2')
    const text = clipboard(w, 'cut')!
    await settle()
    expect(ants(w)).toBe(true)
    await goTo(w, 'C3:D9')
    paste(w, text)
    await settle()
    expect(ops(structure)).toEqual([{ op: 'moveRange', from: { sheet: 'S', range: 'A1:A2' }, to: { sheet: 'S', cell: 'C3' } }])
    expect(address(w)).toBe('C3:C4')
    expect(ants(w)).toBe(false)
    paste(w, text)
    await settle()
    expect(structure).toHaveBeenCalledTimes(1)
    expect(edit).not.toHaveBeenCalled()
  })

  it('other text pastes as text: the clipboard is another’s now, the copy mode ends', async () => {
    const { source, structure, edit } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1:B2')
    clipboard(w, 'copy')
    await goTo(w, 'D5')
    paste(w, 'hello')
    await settle()
    expect(structure).not.toHaveBeenCalled()
    expect(edit).toHaveBeenCalledWith([{ sheet: 'S', cell: 'D5', text: 'hello' }])
    expect(ants(w)).toBe(false)
  })

  it('Esc and a field opened end the copy mode, and a paste of its text does nothing then, as in Excel; a change of cells ends it too', async () => {
    const { source, structure, edit } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1')
    const text = clipboard(w, 'copy')!
    await settle()
    await key(w, 'Escape')
    expect(ants(w)).toBe(false)
    paste(w, text)
    await settle()
    expect(structure).not.toHaveBeenCalled()
    expect(edit).not.toHaveBeenCalled()
    clipboard(w, 'copy')
    await settle()
    await key(w, 'q')
    expect(ants(w)).toBe(false)
    await editorKey(w, 'Escape')
    clipboard(w, 'copy')
    await settle()
    await goTo(w, 'C5')
    await key(w, 'Delete')
    await settle()
    expect(ants(w)).toBe(true)
    await goTo(w, 'C1')
    await key(w, 'Delete')
    await settle()
    expect(ants(w)).toBe(false)
  })

  it('formatting leaves the copy mode on', async () => {
    const { source } = withStructure()
    source.format = vi.fn(async () => OK)
    const w = await mountSheet(source)
    await goTo(w, 'A1:A2')
    clipboard(w, 'copy')
    await key(w, 'b', { ctrlKey: true })
    await settle()
    expect(source.format).toHaveBeenCalled()
    expect(ants(w)).toBe(true)
  })

  it('a refusal shows the source’s message and leaves the selection; a locked target is not sent', async () => {
    const { source, structure } = withStructure(() => ({ ok: false, reason: 'refused', message: 'The copy does not fit' }))
    const w = await mountSheet(source)
    await goTo(w, 'A1:A2')
    const text = clipboard(w, 'copy')!
    await goTo(w, 'D9')
    paste(w, text)
    await settle()
    expect(w.find('.xs-message').text()).toBe('The copy does not fit')
    expect(address(w)).toBe('D9')
    await goTo(w, 'F2')
    paste(w, text)
    await settle()
    expect(structure).toHaveBeenCalledTimes(1)
    expect(w.find('.xs-message').text()).toContain('script')
  })

  it('without structure the copy pastes as its text; a read-only sheet copies, and cuts and pastes nothing', async () => {
    const plain = memorySource(book())
    const edit = vi.spyOn(plain, 'edit')
    const w = await mountSheet(plain)
    await goTo(w, 'A1:B1')
    const text = clipboard(w, 'copy')!
    await goTo(w, 'D5')
    pasteUnread(w)
    await settle()
    expect(edit).toHaveBeenCalledWith([
      { sheet: 'S', cell: 'D5', text: '1' },
      { sheet: 'S', cell: 'E5', text: '10' },
    ])
    document.body.innerHTML = ''
    const { source, structure } = withStructure()
    const ro = await mountSheet(source, { readOnly: true })
    await goTo(ro, 'A1:A2')
    expect(clipboard(ro, 'copy')).toBe('1\r\n2\r\n')
    expect(clipboard(ro, 'cut')).toBeUndefined()
    await goTo(ro, 'D5')
    paste(ro, text)
    await settle()
    expect(structure).not.toHaveBeenCalled()
  })

  it('under the display filter a copy is the shown rows’ text alone, without a mark, after the filter too', async () => {
    const { source, structure, edit } = withStructure()
    const w = await mountSheet(source)
    await w.findAll('.xs-head')[0]!.find('.xs-funnel').trigger('click')
    await settle()
    await w.find('.xs-popover button[title="Sort descending"]').trigger('click')
    await settle()
    await goTo(w, 'A1:A2')
    const text = clipboard(w, 'copy')!
    await settle()
    expect(ants(w)).toBe(false)
    expect(clipboard(w, 'cut')).toBeUndefined()
    // Without the filter the copy is still the text of the rows it showed.
    await w.find('[data-xs-filter-clear]').trigger('click')
    await settle()
    expect(ants(w)).toBe(false)
    await goTo(w, 'D1')
    paste(w, text)
    await settle()
    expect(structure).not.toHaveBeenCalled()
    expect(edit).toHaveBeenCalled()
  })
})

describe('the cells’ menu', () => {
  const items = () => [...document.querySelectorAll<HTMLElement>('[data-xs-menu]')].map((b) => b.dataset.xsMenu)
  const disabled = (id: string) => document.querySelector<HTMLButtonElement>(`[data-xs-menu="${id}"]`)!.disabled
  const pick = async (id: string) => {
    document.querySelector<HTMLElement>(`[data-xs-menu="${id}"]`)!.click()
    await settle()
  }
  async function context(w: VueWrapper, cell: string): Promise<void> {
    await w.find(`[data-cell="${cell}"]`).trigger('contextmenu', { button: 2, clientX: 200, clientY: 100 })
    await settle()
  }
  function clipboardApi(read: () => Promise<string>) {
    const store = { text: '' }
    const api = { writeText: vi.fn(async (t: string) => void (store.text = t)), readText: vi.fn(read) }
    Object.defineProperty(navigator, 'clipboard', { value: api, configurable: true })
    return { api, store }
  }

  it('a right click opens Cut, Copy and Paste over the selection, then the freeze and the precedents; a cell outside it is selected first', async () => {
    const { source } = withStructure()
    const w = await mountSheet(source, { messages: ruMessages })
    await goTo(w, 'A1:B2')
    await context(w, '2:2')
    expect(items()).toEqual(['cut', 'copy', 'paste', 'freeze', 'freezeTopRow', 'freezeFirstColumn', 'precedents'])
    expect(document.querySelector('[data-xs-menu-box]')!.getAttribute('aria-label')).toBe('Ячейки')
    expect(document.querySelector('.xs-menu-caption')!.textContent).toBe('A1:B2')
    expect(document.querySelector('[data-xs-menu="paste"]')!.textContent!.trim()).toBe('Вставить')
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    await settle()
    await context(w, '4:3')
    expect(address(w)).toBe('C4')
  })

  it('Copy writes the text through the clipboard API; Paste reads it back and pastes the range; Cut and Paste move it', async () => {
    const { source, structure } = withStructure()
    const { api, store } = clipboardApi(async () => store.text)
    const w = await mountSheet(source)
    await goTo(w, 'A1:B1')
    await context(w, '1:1')
    await pick('copy')
    expect(api.writeText).toHaveBeenCalledWith('1\t10\r\n')
    expect(ants(w)).toBe(true)
    await context(w, '6:4')
    await pick('paste')
    await context(w, '3:1')
    await pick('cut')
    await context(w, '8:2')
    await pick('paste')
    expect(ops(structure)).toEqual([
      { op: 'copyRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', range: 'D6' } },
      { op: 'moveRange', from: { sheet: 'S', range: 'A3' }, to: { sheet: 'S', cell: 'B8' } },
    ])
    expect(address(w)).toBe('B8')
    store.text = 'outside'
    await context(w, '9:1')
    await pick('paste')
    expect(structure).toHaveBeenCalledTimes(2)
    expect(cellText(w, 9, 1)).toBe('outside')
  })

  it('a browser that will not give the clipboard: Paste pastes only our copy, and without one says how to paste', async () => {
    const { source, structure } = withStructure()
    clipboardApi(async () => Promise.reject(new DOMException('denied', 'NotAllowedError')))
    const w = await mountSheet(source)
    await context(w, '2:1')
    expect(disabled('paste')).toBe(false)
    await pick('paste')
    expect(w.find('.xs-message').text()).toContain('Ctrl+V')
    await context(w, '2:1')
    await pick('copy')
    await context(w, '5:5')
    await pick('paste')
    expect(ops(structure)).toEqual([{ op: 'copyRange', from: { sheet: 'S', range: 'A2' }, to: { sheet: 'S', range: 'E5' } }])
  })

  it('a read-only sheet’s menu only copies and traces; without the clipboard API Paste waits for a copy', async () => {
    const ro = await mountSheet(withStructure().source, { readOnly: true })
    await context(ro, '1:1')
    expect(items()).toEqual(['copy', 'precedents'])
    document.body.innerHTML = ''
    Object.defineProperty(navigator, 'clipboard', { value: undefined, configurable: true })
    const w = await mountSheet(withStructure().source)
    await context(w, '1:1')
    expect(disabled('paste')).toBe(true)
    await pick('copy')
    await context(w, '3:3')
    expect(disabled('paste')).toBe(false)
  })

  it('a finger held on a cell opens it, and the lift keeps the selection', async () => {
    const w = await mountSheet(withStructure().source)
    await goTo(w, 'A1:B2')
    const touch = { pointerType: 'touch', isPrimary: true, button: 0, clientX: 90, clientY: 50 }
    const cell = w.find('[data-cell="2:2"]')
    await cell.trigger('pointerdown', touch)
    await new Promise((r) => setTimeout(r, 560))
    await settle()
    expect(items()).toEqual(['cut', 'copy', 'paste', 'freeze', 'freezeTopRow', 'freezeFirstColumn', 'precedents'])
    await cell.trigger('pointerup', touch)
    await cell.trigger('click')
    await settle()
    expect(address(w)).toBe('A1:B2')
  })
})

describe('with the local source', () => {
  it('a formula copied and pasted moves its references, a cut one keeps them, Control+Z takes each back', async () => {
    const source = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'A1*10' } } }] }, { calc: 'inline' })
    const states: { formula?: string }[] = []
    const w = await mountSheet(source, { onState: (st: { formula?: string } | null) => st && states.push(st) })
    await settle(12)
    await goTo(w, 'A1:B1')
    const text = clipboard(w, 'copy')!
    await goTo(w, 'D5')
    paste(w, text)
    await settle(12)
    expect([cellText(w, 5, 4), cellText(w, 5, 5)]).toEqual(['1', '10'])
    await goTo(w, 'E5')
    expect(states.at(-1)!.formula).toBe('D5*10')
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    // The undo of a copy onto empty cells lists no change for them: `./local` reads the sheet again, so the grid has them empty too.
    expect(((await source.window({ sheet: 'S', row: 5, rows: 1, col: 4, cols: 2 })) as SheetWindow).text).toEqual([['', '']])
    expect([cellText(w, 5, 4), cellText(w, 5, 5)]).toEqual(['', ''])
    await goTo(w, 'B1')
    const cut = clipboard(w, 'cut')!
    await goTo(w, 'D2')
    paste(w, cut)
    await settle(12)
    expect([cellText(w, 1, 2), cellText(w, 2, 4)]).toEqual(['', '10'])
    expect(states.at(-1)!.formula).toBe('A1*10')
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    expect([cellText(w, 1, 2), cellText(w, 2, 4)]).toEqual(['10', ''])
    await source.close()
  })
})

describe('the source’s copy text (SheetSource.copyText)', () => {
  function writable() {
    const api = { writeText: vi.fn(async (_t: string) => {}), readText: vi.fn(async () => '') }
    Object.defineProperty(navigator, 'clipboard', { value: api, configurable: true })
    return api
  }

  it('the copy event takes the shown text at once, the source’s text then goes through the clipboard API; a paste of either is the book’s copy', async () => {
    const { source, structure } = withStructure()
    const copyText = vi.fn(async () => 'EXACT\r\n')
    source.copyText = copyText
    const api = writable()
    const w = await mountSheet(source)
    await goTo(w, 'A1:B1')
    expect(clipboard(w, 'copy')).toBe('1\t10\r\n')
    await settle()
    expect(copyText).toHaveBeenCalledWith('S', 'A1:B1', { trailingNewline: true })
    expect(api.writeText).toHaveBeenCalledWith('EXACT\r\n')
    await goTo(w, 'D5')
    paste(w, 'EXACT\r\n')
    await settle()
    await goTo(w, 'D7')
    paste(w, '1\t10\r\n')
    await settle()
    expect(ops(structure)).toEqual([
      { op: 'copyRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', range: 'D5' } },
      { op: 'copyRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', range: 'D7' } },
    ])
  })

  it('the same text or a refusal writes nothing more; a Mac asks without the last CRLF', async () => {
    vi.spyOn(navigator, 'platform', 'get').mockReturnValue('MacIntel')
    const { source } = withStructure()
    const copyText = vi.fn(async () => '1\t10')
    source.copyText = copyText
    const api = writable()
    const w = await mountSheet(source)
    await goTo(w, 'A1:B1')
    expect(clipboard(w, 'copy')).toBe('1\t10')
    await settle()
    expect(copyText).toHaveBeenCalledWith('S', 'A1:B1', { trailingNewline: false })
    copyText.mockRejectedValueOnce(new Error('too many cells'))
    clipboard(w, 'copy')
    await settle()
    expect(api.writeText).not.toHaveBeenCalled()
  })
})
