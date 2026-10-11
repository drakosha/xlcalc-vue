// Entry through the source's optional methods: Control+Return (`fill`), a paste in one call (`paste`), undo and redo,
// `select` of an outcome, and the text a cell opens with from its window (`editText`).
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import type { CellUpdate, EditOutcome } from '../src/source'
import { active, address, cellText, editorKey, editorOpen, goTo, key, mountSheet, paste, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const book = (): MemoryBook => ({
  sheets: [
    { name: 'S', cells: { A1: 'a1', B1: 'b1', A2: 10, B2: { value: 20, formula: 'A2*2' }, C3: 'c3' }, locked: [{ range: 'D1:D10', reason: 'the script owns it' }] },
    { name: 'T', cells: { A1: 't' } },
  ],
})
const cell = (ref: string, text: string, value: CellUpdate['value'] = text): CellUpdate => ({ sheet: 'S', cell: ref, text, value, formula: null })
/** A call whose answer the test gives when it likes. */
function held() {
  let answer: (o: EditOutcome) => void = () => {}
  const call = vi.fn((..._args: unknown[]) => new Promise<EditOutcome>((resolve) => (answer = resolve)))
  return { call, answer: (o: EditOutcome) => answer(o) }
}
const field = (w: VueWrapper) => w.find('.xs-editor').element as HTMLTextAreaElement

describe('Control+Return', () => {
  it('with `fill` and a range enters the text into the range from the active cell; the selection stays', async () => {
    const source = memorySource(book())
    const fill = held()
    source.fill = fill.call
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await goTo(w, 'B1:B3')
    await key(w, '=')
    await type(w, '=A1*2')
    await editorKey(w, 'Enter', { ctrlKey: true })
    await settle()
    expect(fill.call).toHaveBeenCalledTimes(1)
    expect(fill.call).toHaveBeenCalledWith('S', 'B1:B3', '=A1*2', 'B1')
    expect(edit).not.toHaveBeenCalled()
    expect(editorOpen(w)).toBe(false)
    expect(active(w)).toBe('1:2')
    expect(address(w)).toBe('B1:B3')
    // In flight every cell of the range shows the typed text as pending.
    for (const r of [1, 2, 3]) {
      expect(w.find(`[data-cell="${r}:2"]`).classes()).toContain('xs-pending')
      expect(cellText(w, r, 2)).toBe('=A1*2')
    }
    fill.answer({ ok: true, changes: [cell('B1', '0', 0), cell('B2', '20', 20), cell('B3', '0', 0)] })
    await settle()
    expect(w.find('[data-cell="2:2"]').classes()).not.toContain('xs-pending')
    expect([1, 2, 3].map((r) => cellText(w, r, 2))).toEqual(['0', '20', '0'])
    expect(address(w)).toBe('B1:B3')
  })

  it('Cmd+Return in the formula bar sends the text even when it did not change', async () => {
    const source = memorySource(book())
    const fill = held()
    source.fill = fill.call
    const w = await mountSheet(source)
    await goTo(w, 'B1:C2')
    await w.find('.xs-fx').trigger('click')
    await settle()
    await w.find('.xs-fx-input').trigger('keydown', { key: 'Enter', metaKey: true })
    await settle()
    expect(fill.call).toHaveBeenCalledWith('S', 'B1:C2', 'b1', 'B1')
    expect(address(w)).toBe('B1:C2')
  })

  it('without `fill`, or on one cell, writes the active cell only and the cursor stays', async () => {
    const source = memorySource(book())
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await goTo(w, 'B1:B3')
    await key(w, 'q')
    await editorKey(w, 'Enter', { ctrlKey: true })
    await settle()
    expect(edit.mock.calls).toEqual([[[{ sheet: 'S', cell: 'B1', text: 'q' }]]])
    expect(active(w)).toBe('1:2')
    expect(address(w)).toBe('B1:B3')
    expect(cellText(w, 2, 2)).toBe('20')

    const withFill = memorySource(book())
    const fill = held()
    withFill.fill = fill.call
    const one = vi.spyOn(withFill, 'edit')
    const v = await mountSheet(withFill)
    await goTo(v, 'C2')
    await key(v, 'z')
    await editorKey(v, 'Enter', { ctrlKey: true })
    await settle()
    expect(fill.call).not.toHaveBeenCalled()
    expect(one.mock.calls).toEqual([[[{ sheet: 'S', cell: 'C2', text: 'z' }]]])
    expect(active(v)).toBe('2:3')
  })

  it('checks locks over the whole range, and `invalid` opens the field again over the range', async () => {
    const source = memorySource(book())
    const fill = held()
    source.fill = fill.call
    const w = await mountSheet(source)
    await goTo(w, 'C1:D2')
    await key(w, 'x')
    await editorKey(w, 'Enter', { ctrlKey: true })
    await settle()
    expect(fill.call).not.toHaveBeenCalled()
    expect(w.find('.xs-message').text()).toContain('the script owns it')
    expect(editorOpen(w)).toBe(true)
    await editorKey(w, 'Escape')

    await goTo(w, 'A4:B5')
    await key(w, '=')
    await type(w, '=SUM(A1')
    await editorKey(w, 'Enter', { ctrlKey: true })
    fill.answer({ ok: false, reason: 'invalid', message: 'Missing bracket', at: 4 })
    await settle()
    expect(editorOpen(w)).toBe(true)
    expect(field(w).value).toBe('=SUM(A1')
    expect(field(w).selectionStart).toBe(4)
    expect(address(w)).toBe('A4:B5')
    expect(w.find('.xs-message').text()).toBe('Missing bracket')
  })
})

describe('paste through the source', () => {
  it('goes in one call at the top-left cell of the selection; the selection goes to `select`', async () => {
    const source = memorySource(book())
    const call = vi.fn(async (..._args: unknown[]): Promise<EditOutcome> => ({
      ok: true,
      changes: [cell('B5', '1', 1), cell('C5', '2', 2), cell('B6', '3', 3)],
      select: { sheet: 'S', range: 'B5:C7' },
    }))
    source.paste = call
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await goTo(w, 'C6')
    await key(w, 'ArrowUp', { shiftKey: true })
    await key(w, 'ArrowLeft', { shiftKey: true })
    expect(address(w)).toBe('B5:C6')
    paste(w, '1\t2\n3\t\n')
    await settle()
    expect(call).toHaveBeenCalledTimes(1)
    expect(call).toHaveBeenCalledWith('S', 'B5', '1\t2\n3\t\n')
    expect(edit).not.toHaveBeenCalled()
    expect(cellText(w, 6, 2)).toBe('3')
    expect(address(w)).toBe('B5:C7')
    expect(active(w)).toBe('5:2')
  })

  it('checks locks over the rectangle the text makes', async () => {
    const source = memorySource(book())
    const call = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [] }))
    source.paste = call
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    paste(w, '1\t2\t3')
    await settle()
    expect(call).not.toHaveBeenCalled()
    expect(w.find('.xs-message').text()).toContain('the script owns it')
  })

  it('an edit answered with `select` moves the selection there', async () => {
    const source = memorySource(book())
    const edit = source.edit!.bind(source)
    source.edit = async (entries) => ({ ...(await edit(entries)), select: { sheet: 'S', range: 'A1:B2' } })
    const w = await mountSheet(source)
    await key(w, 'k')
    await editorKey(w, 'Enter')
    await settle()
    expect(cellText(w, 1, 1)).toBe('k')
    expect(address(w)).toBe('A1:B2')
  })
})

describe('undo and redo', () => {
  it('Control+Z, Control+Y, Cmd+Shift+Z and a Russian layout call the source; changes apply, the selection goes to `select`', async () => {
    const source = memorySource(book())
    const undo = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [cell('A1', 'old')], select: { sheet: 'S', range: 'A1:B2' } }))
    const redo = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [cell('A1', 'new')], select: { sheet: 'S', range: 'C3' } }))
    source.undo = undo
    source.redo = redo
    const w = await mountSheet(source)
    await goTo(w, 'E5')
    await key(w, 'z', { code: 'KeyZ', ctrlKey: true })
    await settle()
    expect(undo).toHaveBeenCalledTimes(1)
    expect(cellText(w, 1, 1)).toBe('old')
    expect(address(w)).toBe('A1:B2')
    await key(w, 'y', { code: 'KeyY', ctrlKey: true })
    await settle()
    expect(redo).toHaveBeenCalledTimes(1)
    expect(cellText(w, 1, 1)).toBe('new')
    expect(address(w)).toBe('C3')
    await key(w, 'я', { code: 'KeyZ', ctrlKey: true })
    await key(w, 'Z', { code: 'KeyZ', metaKey: true, shiftKey: true })
    await key(w, 'н', { code: 'KeyY', metaKey: true })
    await settle()
    expect(undo).toHaveBeenCalledTimes(2)
    expect(redo).toHaveBeenCalledTimes(3)
    // In the field Control+Z is the field's own undo.
    await key(w, 'x')
    await editorKey(w, 'z', { code: 'KeyZ', ctrlKey: true })
    await settle()
    expect(undo).toHaveBeenCalledTimes(2)
    expect(editorOpen(w)).toBe(true)
  })

  it('sends one call at a time, and goes to another sheet named by `select`', async () => {
    const source = memorySource(book())
    const undo = held()
    source.undo = undo.call
    const w = await mountSheet(source)
    await key(w, 'z', { code: 'KeyZ', ctrlKey: true })
    await key(w, 'z', { code: 'KeyZ', ctrlKey: true })
    await settle()
    expect(undo.call).toHaveBeenCalledTimes(1)
    undo.answer({ ok: true, changes: [{ sheet: 'T', cell: 'C4', text: 'back', value: 'back', formula: null }], select: { sheet: 'T', range: 'C4' } })
    await settle()
    expect(w.find('.xs-tab-on').text()).toBe('T')
    expect(active(w)).toBe('4:3')
    await key(w, 'z', { code: 'KeyZ', ctrlKey: true })
    await settle()
    expect(undo.call).toHaveBeenCalledTimes(2)
  })

  it('leaves the selection where the person moved it while the call was out', async () => {
    const source = memorySource(book())
    const undo = held()
    source.undo = undo.call
    const w = await mountSheet(source)
    await key(w, 'z', { code: 'KeyZ', ctrlKey: true })
    await goTo(w, 'C2')
    undo.answer({ ok: true, changes: [cell('A1', 'old')], select: { sheet: 'S', range: 'A1' } })
    await settle()
    expect(cellText(w, 1, 1)).toBe('old')
    expect(address(w)).toBe('C2')
  })
})

describe('the text a cell opens with', () => {
  /** A memory book whose windows carry `editText` from `texts` (cells left out open empty). */
  function withEditText(texts: Record<string, string>) {
    const source = memorySource({
      sheets: [{ name: 'S', cells: { A1: { value: 46096, text: '15 марта 2026 г.' }, A2: { value: 0.05, text: '5%' }, A3: { value: 'x', text: 'x' }, A4: { value: 3, formula: 'A2+1' } } }],
    })
    const window = source.window.bind(source)
    source.window = async (req, signal) => {
      const got = await window(req, signal)
      return 'pending' in got ? got : { ...got, editText: { ...texts } }
    }
    return source
  }
  const open = async (w: VueWrapper, place: string) => {
    await goTo(w, place)
    await key(w, 'F2')
    await settle()
    const text = field(w).value
    await editorKey(w, 'Escape')
    return text
  }
  const bar = (w: VueWrapper) => w.find('.xs-fx-text').text()

  it("opens with the window's editText, empty for a cell left out; the formula bar shows the same", async () => {
    const source = withEditText({ A1: '15.03.2026', A2: '5%', A4: '=A2+1' })
    const w = await mountSheet(source)
    expect(await open(w, 'A1')).toBe('15.03.2026')
    expect(bar(w)).toBe('15.03.2026')
    expect(await open(w, 'A2')).toBe('5%')
    expect(await open(w, 'A3')).toBe('')
    expect(bar(w)).toBe('')
    expect(await open(w, 'A4')).toBe('=A2+1')
    expect(bar(w)).toBe('=A2+1')
  })

  it('takes the new text from a change; without the dictionary the component writes it as before', async () => {
    const source = withEditText({ A1: '15.03.2026' })
    const w = await mountSheet(source)
    source.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: 'A1', text: '16 марта 2026 г.', value: 46097, formula: null, editText: '16.03.2026' }] })
    await settle()
    expect(cellText(w, 1, 1)).toBe('16 марта 2026 г.')
    expect(await open(w, 'A1')).toBe('16.03.2026')

    const plain = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: { value: 46096, text: '15 марта 2026 г.' }, A3: 'x' } }] }))
    expect(await open(plain, 'A1')).toBe('46096')
    expect(await open(plain, 'A3')).toBe('x')
  })
})
