// Underline (single and double) and strike reach the element that carries the text in every kind of cell: decorations
// do not pass to an absolutely positioned overflow run, so they must sit on the text itself.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { mountSheet } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

/** The innermost element of a cell's content box holding its text. */
function carrier(w: VueWrapper, r: number, c: number, text: string): HTMLElement {
  const found = w
    .find(`[data-cell="${r}:${c}"]`)
    .findAll('.xs-in *')
    .filter((e) => e.element.children.length === 0 && e.text().includes(text))
  expect(found.length, `text carrier of ${r}:${c}`).toBe(1)
  return found[0]!.element as HTMLElement
}

describe('text decoration', () => {
  it('sits on the text in plain, overflow, centerContinuous, fill, link and merged cells', async () => {
    const book: MemoryBook = {
      styles: [
        {},
        { font: { underline: 'single' } },
        { font: { underline: 'double' } },
        { font: { strike: true } },
        { font: { underline: 'single', strike: true }, alignment: { horizontal: 'centerContinuous' } },
        { font: { strike: true }, alignment: { horizontal: 'fill' } },
        { alignment: { horizontal: 'centerContinuous' } },
      ],
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: 'overflow', style: 1 },
            A2: { value: 'plain', style: 2 },
            B2: 'busy',
            A3: { value: 'struck', style: 3 },
            B3: 'busy',
            A4: { value: 'across', style: 4 },
            B4: { style: 6 },
            A5: { value: '-', style: 5 },
            A6: { value: 'link', style: 3, link: { href: 'https://example.org' } },
            A7: { value: 'merged', style: 2 },
          },
          merges: ['A7:B8'],
        },
      ],
    }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    const lines = (r: number, c: number, text: string) => carrier(w, r, c, text).style.textDecorationLine
    expect(w.find('[data-cell="1:1"] .xs-ov').exists()).toBe(true)
    expect(lines(1, 1, 'overflow')).toBe('underline')
    expect(lines(2, 1, 'plain')).toBe('underline')
    expect(carrier(w, 2, 1, 'plain').style.textDecorationStyle).toBe('double')
    expect(lines(3, 1, 'struck')).toBe('line-through')
    expect(w.find('[data-cell="4:1"] .xs-ov').exists()).toBe(true)
    expect(lines(4, 1, 'across')).toBe('underline line-through')
    expect(lines(5, 1, '-')).toBe('line-through')
    // A link keeps its own underline and adds the style's strike.
    expect(lines(6, 1, 'link')).toBe('underline line-through')
    expect(lines(7, 1, 'merged')).toBe('underline')
  })
})
