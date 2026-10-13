// Excel's Point mode: while a formula is edited with the caret where a reference may stand, a click on the sheet puts the
// cell's address at the caret, a drag a range's, a letter or a number a whole column's or row's; the next click replaces it.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { refFits } from '../src/core/point'
import { active, editorKey, editorOpen, key, mountSheet, settle, type } from './helpers'

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

/** The middle of a cell on the page: the scroller is at 0, 0; letters 24 px high, numbers 56 px wide, rows 20 px, columns 62 px. */
const at = (row: number, col: number) => ({ clientX: 56 + (col - 1) * 62 + 31, clientY: 24 + (row - 1) * 20 + 10 })
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

describe('where a reference may stand', () => {
  it('after =, an operator, (, : and the list separator; not after a value or in a string', () => {
    for (const t of ['=', '=1+', '=SUM(', '=SUM( ', '=A1:', '=SUM(A1,', '=A1>=', '=-', '=@']) expect([t, refFits(t, t.length, ',')]).toEqual([t, true])
    for (const t of ['=5', '=A1', '=SUM(A1)', '="a+', "='a (", '=T[', 'abc+', '=1%']) expect([t, refFits(t, t.length, ',')]).toEqual([t, false])
    // ru-RU: `;` separates arguments, `,` is the decimal point
    expect(refFits('=СУММ(1;', 8, ';')).toBe(true)
    expect(refFits('=СУММ(1,', 8, ';')).toBe(false)
    // a selection's start counts: `A1` selected in `=A1+B1`
    expect(refFits('=A1+B1', 1, ',')).toBe(true)
  })
})

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
    await letter.trigger('pointerdown', { button: 0, pointerId: 1, pointerType: 'mouse', clientX: 56 + 62 + 31, clientY: 12 })
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
    for (const text of ['=5', 'abc']) {
      const w = await mountSheet(recorded().source)
      await typeIn(w, text)
      const down = await press(w, 3, 3)
      expect(down.defaultPrevented).toBe(false)
      expect(field(w).value).not.toContain('C3')
      expect(active(w)).toBe('3:3')
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
    await w.find('.xs-fx').trigger('click')
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
