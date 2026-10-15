// Excel's Point mode: while a formula is edited with the caret where a reference may stand, a click on the sheet puts the
// cell's address at the caret, a drag a range's, a letter or a number a whole column's or row's; the next click replaces it.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { colPx } from '../src/core/geometry'
import { active, editorKey, editorOpen, goTo, key, mountSheet, pressBar, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'a1', B2: 5, C3: 7 } }] })

/** A memory source that records the edits it gets. */
function recorded() {
  const source = memorySource(book())
  const edits: Parameters<NonNullable<typeof source.edit>>[0][] = []
  const edit = source.edit!.bind(source)
  source.edit = (entries) => {
    edits.push(entries)
    return edit(entries)
  }
  return { source, edits }
}

/** A column without a width, in pixels. */
const COL = colPx(null)
/** The middle of a cell on the page: the scroller is at 0, 0; letters 24 px high, numbers 56 px wide, rows 20 px, columns `COL`. */
const at = (row: number, col: number) => ({ clientX: 56 + (col - 1) * COL + COL / 2, clientY: 24 + (row - 1) * 20 + 10 })
const field = (w: VueWrapper) => w.find('.xs-editor').element as HTMLTextAreaElement
/** Open the field on the active cell with `text` typed, the caret at its end. */
async function typeIn(w: VueWrapper, text: string): Promise<void> {
  await key(w, text[0]!)
  await type(w, text)
  field(w).setSelectionRange(text.length, text.length)
}
/** A mouse press on a cell, its mousedown, a drag through cells and the release, as the browser sends them. */
async function press(w: VueWrapper, row: number, col: number, opts: { shiftKey?: boolean; through?: [number, number][] } = {}): Promise<MouseEvent> {
  const td = w.find(`[data-cell="${row}:${col}"]`)
  await td.trigger('pointerdown', { button: 0, pointerId: 1, pointerType: 'mouse', shiftKey: !!opts.shiftKey, ...at(row, col) })
  const down = new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })
  td.element.dispatchEvent(down)
  for (const [r, c] of opts.through ?? []) await w.find('.xs-scroller').trigger('pointermove', { pointerId: 1, pointerType: 'mouse', ...at(r, c) })
  await w.find('.xs-scroller').trigger('pointerup', { pointerId: 1, pointerType: 'mouse' })
  await td.trigger('click', { button: 0, shiftKey: !!opts.shiftKey })
  await settle()
  return down
}

describe('Point mode', () => {
  it('a click after = puts the cell at the caret; the field stays open, the active cell stays, the focus stays', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await typeIn(w, '=')
    const down = await press(w, 3, 3)
    expect(down.defaultPrevented).toBe(true)
    expect(editorOpen(w)).toBe(true)
    expect(field(w).value).toBe('=C3')
    expect(field(w).selectionStart).toBe(3)
    expect(active(w)).toBe('1:1')
    expect(edits).toEqual([])
    await editorKey(w, 'Enter')
    await settle()
    expect(edits).toEqual([[{ sheet: 'S', cell: 'A1', text: '=C3' }]])
  })

  it('a drag gives the range; the next click replaces it, Shift extends from its first corner', async () => {
    const w = await mountSheet(recorded().source)
    await typeIn(w, '=SUM(')
    await press(w, 2, 2, { through: [[3, 3], [4, 4]] })
    expect(field(w).value).toBe('=SUM(B2:D4')
    await press(w, 5, 1)
    expect(field(w).value).toBe('=SUM(A5')
    await press(w, 6, 3, { shiftKey: true })
    expect(field(w).value).toBe('=SUM(A5:C6')
    expect(active(w)).toBe('1:1')
  })

  it('after an operator typed a click adds a reference; the text after the caret stays', async () => {
    const w = await mountSheet(recorded().source)
    await typeIn(w, '=')
    await press(w, 2, 2)
    await type(w, '=B2+)')
    field(w).setSelectionRange(4, 4)
    await press(w, 3, 3)
    expect(field(w).value).toBe('=B2+C3)')
    expect(field(w).selectionStart).toBe(6)
  })

  it('selected text where a reference may stand is replaced', async () => {
    const w = await mountSheet(recorded().source)
    await typeIn(w, '=A1+B1')
    field(w).setSelectionRange(1, 3)
    await press(w, 4, 4)
    expect(field(w).value).toBe('=D4+B1')
  })

  it('a letter puts its column, a number its row', async () => {
    const w = await mountSheet(recorded().source)
    await typeIn(w, '=SUM(')
    const letter = w.find('th[data-col="2"]')
    await letter.trigger('pointerdown', { button: 0, pointerId: 1, pointerType: 'mouse', clientX: 56 + COL + COL / 2, clientY: 12 })
    await letter.trigger('click', { button: 0 })
    await settle()
    expect(field(w).value).toBe('=SUM(B:B')
    expect(editorOpen(w)).toBe(true)
    await type(w, '=SUM(B:B,')
    field(w).setSelectionRange(9, 9)
    const number = w.find('th[data-row="3"]')
    await number.trigger('pointerdown', { button: 0, pointerId: 1, pointerType: 'mouse', clientX: 20, clientY: 24 + 2 * 20 + 10 })
    await number.trigger('click', { button: 0 })
    await settle()
    expect(field(w).value).toBe('=SUM(B:B,3:3')
    expect(active(w)).toBe('1:1')
  })

  it('a double click while pointing does not open another cell', async () => {
    const w = await mountSheet(recorded().source)
    await typeIn(w, '=')
    await press(w, 3, 3)
    await w.find('[data-cell="3:3"]').trigger('dblclick')
    await settle()
    expect(field(w).value).toBe('=C3')
    expect(active(w)).toBe('1:1')
  })

  it('where no reference may stand the click is the sheet’s as before: no address goes in, the cell is picked', async () => {
    for (const text of ['=5', 'abc', '=A1)', '="a ']) {
      const w = await mountSheet(recorded().source)
      await typeIn(w, text)
      const down = await press(w, 3, 3)
      expect(down.defaultPrevented).toBe(false)
      expect(field(w).value).not.toContain('C3')
      expect(active(w)).toBe('3:3')
      w.unmount()
    }
  })

  it('a click puts the cell where an arrow would point: after a leading + or -, a space after an operand, a point', async () => {
    for (const [text, got] of [['+', '+C3'], ['-', '-C3'], ['=A1 ', '=A1 C3'], ['=1.', '=1.C3'], ['=S!', '=S!C3']]) {
      const w = await mountSheet(recorded().source)
      await typeIn(w, text!)
      await press(w, 3, 3)
      expect([text, field(w).value, active(w)]).toEqual([text, got, '1:1'])
      w.unmount()
    }
  })

  it('the list separator of the book’s locale: `;` in ru-RU, a comma there is a decimal point', async () => {
    const w = await mountSheet(recorded().source, { locale: 'ru-RU' })
    await settle()
    await typeIn(w, '=СУММ(1;')
    await press(w, 2, 2)
    expect(field(w).value).toBe('=СУММ(1;B2')
  })

  it('works from the formula bar too', async () => {
    const w = await mountSheet(recorded().source)
    await pressBar(w, 0)
    await settle()
    const bar = w.find('.xs-fx textarea').element as HTMLTextAreaElement
    bar.value = '=1+'
    await w.find('.xs-fx textarea').trigger('input')
    bar.setSelectionRange(3, 3)
    await press(w, 2, 2)
    expect(bar.value).toBe('=1+B2')
    expect(active(w)).toBe('1:1')
  })
})

describe('Point mode with the arrows', () => {
  /** An arrow in the field, the caret where it stood. */
  async function arrow(w: VueWrapper, k: string, shiftKey = false): Promise<void> {
    await editorKey(w, k, { shiftKey })
    await settle()
  }

  it('after = typed over a cell an arrow puts the cell beside it; arrows move the reference, Shift extends it, a typed operator fixes it', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    await typeIn(w, '=')
    await arrow(w, 'ArrowRight')
    expect(field(w).value).toBe('=C2')
    expect(active(w)).toBe('2:2')
    await arrow(w, 'ArrowRight')
    expect(field(w).value).toBe('=D2')
    await arrow(w, 'ArrowDown', true)
    await arrow(w, 'ArrowDown', true)
    expect(field(w).value).toBe('=D2:D4')
    // a plain arrow moves from the reference's first cell and drops the range
    await arrow(w, 'ArrowLeft')
    expect(field(w).value).toBe('=C2')
    await type(w, '=C2+')
    field(w).setSelectionRange(4, 4)
    // the next reference starts from the edited cell again
    await arrow(w, 'ArrowUp')
    expect(field(w).value).toBe('=C2+B1')
    expect(field(w).selectionStart).toBe(6)
    await editorKey(w, 'Enter')
    await settle()
    expect(edits).toEqual([[{ sheet: 'S', cell: 'B2', text: '=C2+B1' }]])
  })

  it('at the sheet’s edge the reference stays', async () => {
    const w = await mountSheet(recorded().source)
    await typeIn(w, '=')
    await arrow(w, 'ArrowUp')
    expect(field(w).value).toBe('=A1')
    await arrow(w, 'ArrowLeft')
    expect(field(w).value).toBe('=A1')
  })

  it('a reference the mouse put goes on with the arrows', async () => {
    const w = await mountSheet(recorded().source)
    await typeIn(w, '=')
    await press(w, 3, 3)
    await arrow(w, 'ArrowDown')
    expect(field(w).value).toBe('=C4')
  })

  it('the arrows move the caret in a field F2 opened; F2 turns them to pointing', async () => {
    const w = await mountSheet(recorded().source)
    await goTo(w, 'B2')
    await key(w, 'F2')
    await type(w, '=1+')
    field(w).setSelectionRange(3, 3)
    await arrow(w, 'ArrowRight')
    expect(field(w).value).toBe('=1+')
    // F2 turns the field to Enter mode, where the arrows point
    await editorKey(w, 'F2')
    await arrow(w, 'ArrowRight')
    expect(field(w).value).toBe('=1+C2')
  })

  it('the arrows move the caret in the formula bar', async () => {
    const w = await mountSheet(recorded().source)
    await pressBar(w, 0)
    await settle()
    const bar = w.find('.xs-fx textarea')
    ;(bar.element as HTMLTextAreaElement).value = '=1+'
    await bar.trigger('input')
    ;(bar.element as HTMLTextAreaElement).setSelectionRange(3, 3)
    await bar.trigger('keydown', { key: 'ArrowRight' })
    await settle()
    expect((bar.element as HTMLTextAreaElement).value).toBe('=1+')
  })

  it('the list separator of a ru-RU book', async () => {
    const w = await mountSheet(recorded().source, { locale: 'ru-RU' })
    await settle()
    await goTo(w, 'B2')
    await typeIn(w, '=СУММ(1;')
    await arrow(w, 'ArrowDown')
    expect(field(w).value).toBe('=СУММ(1;B3')
  })
})

describe('Point mode: what an arrow does with nothing pointed, as Excel measured (refAt)', () => {
  /** A key in the field as the browser sends it; the event tells whether the field's own move was stopped. */
  async function press(w: VueWrapper, k: string, init: KeyboardEventInit = {}): Promise<KeyboardEvent> {
    const e = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init })
    field(w).dispatchEvent(e)
    await settle()
    return e
  }

  it('after a whole operand the arrow enters the formula and moves the selection, as Enter does that way', async () => {
    for (const [text, k, to] of [['=5', 'ArrowRight', '2:3'], ['=A1', 'ArrowDown', '3:2'], ['=SUM(A1)', 'ArrowLeft', '2:1'], ['=D3+1', 'ArrowUp', '1:2'], ['=TRUE', 'ArrowRight', '2:3']] as const) {
      const { source, edits } = recorded()
      const w = await mountSheet(source)
      await goTo(w, 'B2')
      await typeIn(w, text)
      await press(w, k)
      expect([text, editorOpen(w), active(w), edits]).toEqual([text, false, to, [[{ sheet: 'S', cell: 'B2', text }]]])
      w.unmount()
    }
  })

  it('with the caret before the formula the arrow puts the cell with an = of its own, as Excel; a value is entered', async () => {
    for (const [text, k, got] of [['=A1', 'ArrowRight', '=D3=A1'], ['+1', 'ArrowDown', '=C4+1']] as const) {
      const w = await mountSheet(recorded().source)
      await goTo(w, 'C3')
      await typeIn(w, text)
      field(w).setSelectionRange(0, 0)
      await press(w, k)
      expect([text, field(w).value, editorOpen(w), active(w)]).toEqual([text, got, true, '3:3'])
      // the reference goes on with the arrows, the = staying
      await press(w, 'ArrowRight')
      expect(field(w).value).toBe(got.replace(/^=([A-Z]+)(\d+)/, (_, c: string, r: string) => `=${String.fromCharCode(c.charCodeAt(0) + 1)}${r}`))
      w.unmount()
    }
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await goTo(w, 'C3')
    await typeIn(w, '5')
    field(w).setSelectionRange(0, 0)
    await press(w, 'ArrowRight')
    expect([editorOpen(w), active(w), edits]).toEqual([false, '3:4', [[{ sheet: 'S', cell: 'C3', text: '5' }]]])
  })

  it('a leading + or - begins a formula: the arrow points', async () => {
    for (const [text, k, got] of [['+', 'ArrowRight', '+D3'], ['-', 'ArrowDown', '-C4'], ['=-', 'ArrowLeft', '=-B3']] as const) {
      const w = await mountSheet(recorded().source)
      await goTo(w, 'C3')
      await typeIn(w, text)
      await press(w, k)
      expect([text, field(w).value, editorOpen(w), active(w)]).toEqual([text, got, true, '3:3'])
      w.unmount()
    }
  })

  it('a space after an operand (the intersection), a point, a sheet’s ! take a reference', async () => {
    for (const [text, got] of [['=A1 ', '=A1 D3'], ['=1.', '=1.D3'], ['=A1.', '=A1.D3'], ['=S!', '=S!D3'], ['=A1:', '=A1:D3']] as const) {
      const w = await mountSheet(recorded().source)
      await goTo(w, 'C3')
      await typeIn(w, text)
      await press(w, 'ArrowRight')
      expect([text, field(w).value, active(w)]).toEqual([text, got, '3:3'])
      w.unmount()
    }
  })

  it('where the formula does not read, or inside a string, a quoted sheet’s name or a table’s brackets, the arrow does nothing', async () => {
    for (const text of ['=A1)', '=1e', '={1', '=$', '="a', "='a b", '=T[']) {
      const { source, edits } = recorded()
      const w = await mountSheet(source)
      await goTo(w, 'C3')
      await typeIn(w, text)
      const e = await press(w, 'ArrowRight')
      expect([text, e.defaultPrevented, field(w).value, editorOpen(w), active(w), edits]).toEqual([text, true, text, true, '3:3', []])
      w.unmount()
    }
  })

  it('in a ru-RU book the locale decides: `=1,5` is entered, `=1.5` and `=A1,` do nothing', async () => {
    for (const [text, open] of [['=1,5', false], ['=1.5', true], ['=A1,', true]] as const) {
      const w = await mountSheet(recorded().source, { locale: 'ru-RU' })
      await settle()
      await goTo(w, 'C3')
      await typeIn(w, text)
      await press(w, 'ArrowRight')
      expect([text, editorOpen(w)]).toEqual([text, open])
      w.unmount()
    }
  })

  it('the first Shift and arrow after = makes a range from the edited cell', async () => {
    const w = await mountSheet(recorded().source)
    await goTo(w, 'C3')
    await typeIn(w, '=')
    await press(w, 'ArrowRight', { shiftKey: true })
    expect(field(w).value).toBe('=C3:D3')
    await press(w, 'ArrowDown', { shiftKey: true })
    expect(field(w).value).toBe('=C3:D4')
    expect(active(w)).toBe('3:3')
  })

  it('Command and an arrow (Control on Windows) jump to the data’s edge by the source’s end, Shift extends; Home, Page Down and Page Up point', async () => {
    const src = memorySource(book())
    const end = vi.fn(async (_sheet: string, cell: string, dir: string) => (cell === 'D3' && dir === 'down' ? 'D6' : cell))
    const w = await mountSheet({ ...src, end })
    await goTo(w, 'C3')
    await typeIn(w, '=')
    await press(w, 'ArrowRight')
    const meta = await press(w, 'ArrowDown', { metaKey: true })
    expect(meta.defaultPrevented).toBe(true)
    expect(end).toHaveBeenLastCalledWith('S', 'D3', 'down')
    expect(field(w).value).toBe('=D6')
    expect(active(w)).toBe('3:3')
    await press(w, 'Escape')
    await typeIn(w, '=')
    await press(w, 'ArrowRight')
    await press(w, 'ArrowDown', { ctrlKey: true, shiftKey: true })
    expect(field(w).value).toBe('=D3:D6')
    await press(w, 'Home')
    expect(field(w).value).toBe('=A3')
    await press(w, 'PageDown')
    const row = Number(/^=A(\d+)$/.exec(field(w).value)?.[1])
    expect(row).toBeGreaterThan(10)
    await press(w, 'PageUp')
    expect(field(w).value).toBe('=A3')
    expect(active(w)).toBe('3:3')
  })
})
