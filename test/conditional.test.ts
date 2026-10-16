// Conditional formats as xlcalc computes them (the window's `appearance`): a rule's style over the cell's, data bars,
// icons of the sets and values hidden; the local source gives them and reads them again after an edit.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { createInlineCalc, type AsyncSession, type BookInput, type Calc } from 'xlcalc'
import { memorySource, type MemoryBook } from '../src'
import { localSource } from '../src/local'
import type { SourceEvent } from '../src/source'
import { editorKey, goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const td = (w: VueWrapper, r: number, c: number) => w.find(`[data-cell="${r}:${c}"]`)
const css = (w: VueWrapper, r: number, c: number) => getComputedStyle(td(w, r, c).element)
const side = (w: VueWrapper, r: number, c: number, s: 'Right' | 'Bottom') => {
  const st = css(w, r, c) as unknown as Record<string, string>
  return [st[`border${s}Width`], st[`border${s}Style`], st[`border${s}Color`]].join(' ')
}
/** The element that carries a cell's text. */
const textEl = (w: VueWrapper, r: number, c: number) => td(w, r, c).find('.xs-in > span')

describe('conditional formats: a rule style', () => {
  it('lays a rule style over the cell own: a field the rule sets wins, the rest shows through', async () => {
    const book: MemoryBook = {
      styles: [{}, { font: { bold: true, color: '#7030A0', underline: 'single' }, fill: { color: '#FFFF00' }, border: { bottom: { style: 'thin', color: '#BFBFBF' }, right: { style: 'thin', color: '#00FF00' } } }],
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: 12, style: 1, appearance: { style: { font: { color: '#006100', bold: false, italic: true, strike: true }, fill: { color: '#C6EFCE' }, border: { right: { style: 'thin', color: '#2F5597' } } } } },
            A2: 1,
            B2: { value: 2, appearance: { style: { fill: { color: '#FFC7CE' } } } },
          },
        },
      ],
    }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    const a1 = css(w, 1, 1)
    expect(a1.color).toBe('#006100')
    expect(a1.fontWeight).not.toBe('700')
    expect(a1.fontStyle).toBe('italic')
    expect(a1.backgroundColor).toBe('#C6EFCE')
    expect(side(w, 1, 1, 'Right')).toBe('1px solid #2F5597')
    expect(side(w, 1, 1, 'Bottom')).toBe('1px solid #BFBFBF')
    expect(getComputedStyle(textEl(w, 1, 1).element).textDecorationLine).toBe('underline line-through')
    // A neighbour's fill by a rule hides the grid line between, as the cell's own fill does.
    expect(side(w, 2, 1, 'Right')).toBe('1px solid #FFC7CE')
  })
})

describe('conditional formats: data bars', () => {
  const book = (): MemoryBook => ({
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 10, appearance: { bar: { start: 0, end: 0.5, color: '#638EC6', gradient: true, border: '#638EC6' } } },
          A2: { value: -10, appearance: { bar: { start: 0.2, end: 0.32, color: '#FF0000', gradient: true, axis: 0.32, axisColor: '#000000' } } },
          A3: { value: 5, appearance: { bar: { start: 0, end: 1, color: '#FFB628', gradient: false } } },
          A4: { value: 4, appearance: { bar: { start: 0.25, end: 0.4, color: '#D6007B', gradient: true, rightToLeft: true } } },
        },
      },
    ],
  })
  const layers = (w: VueWrapper, r: number) => {
    const s = css(w, r, 1)
    return { image: s.backgroundImage, size: s.backgroundSize, pos: s.backgroundPosition }
  }

  it('draws a gradient bar from its start to its end with a border, dark at its start', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    const a1 = layers(w, 1)
    // the fill fades to near white at its end, over a rectangle of the border's colour
    expect(a1.image).toContain('linear-gradient(to right, transparent calc(0% + 1px), #638EC6 calc(0% + 1px), color-mix(in srgb, #638EC6 18%, white) calc(50% - 1px), transparent calc(50% - 1px))')
    expect(a1.image).toContain('linear-gradient(to right, transparent 0%, #638EC6 0%, #638EC6 50%, transparent 50%)')
    expect(a1.size).toBe('calc(100% - 2px) calc(100% - 6px), calc(100% - 2px) calc(100% - 4px)')
    expect(a1.pos).toBe('left 1px top 2px, left 1px top 1px')
    expect(td(w, 1, 1).text()).toBe('10')
  })

  it('draws a negative bar dark at its axis and the axis dashed across the cell', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    const a2 = layers(w, 2)
    expect(a2.image).toContain('repeating-linear-gradient(to bottom, #000000 0px 3px, transparent 3px 4px)')
    expect(a2.image).toContain('linear-gradient(to right, transparent 20%, color-mix(in srgb, #FF0000 18%, white) 20%, #FF0000 32%, transparent 32%)')
    expect(a2.size).toBe('1px 100%, calc(100% - 2px) calc(100% - 4px)')
    // the axis's left edge at 32% of the bar's room: 1px + 0.32 × (width − 2px)
    expect(a2.pos).toBe('calc(32% + 0.68px) 0px, left 1px top 1px')
  })

  it('draws a solid bar of one colour and a bar from the right', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    expect(layers(w, 3).image).toBe('linear-gradient(to right, transparent 0%, #FFB628 0%, #FFB628 100%, transparent 100%)')
    expect(layers(w, 4).image).toBe('linear-gradient(to left, transparent 25%, #D6007B 25%, color-mix(in srgb, #D6007B 18%, white) 40%, transparent 40%)')
  })
})

describe('conditional formats: icons and hidden values', () => {
  it('draws the icon of its set and number at the left of the cell, the value after it', async () => {
    const book: MemoryBook = {
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: 97, appearance: { icon: { set: '3Arrows', index: 2 } } },
            A2: { value: 3, appearance: { icon: { set: '3Arrows', index: 0 } } },
            A3: { value: 4, appearance: { icon: { set: '5Quarters', index: 4 } } },
          },
        },
      ],
    }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    const icon = (r: number) => td(w, r, 1).find('.xs-in > svg.xs-cf-icon')
    expect(icon(1).attributes('data-icon')).toBe('3Arrows:2')
    // Excel's green arrow up and red arrow down
    expect(icon(1).html()).toContain('#66C07E')
    expect(icon(2).html()).toContain('#F6888A')
    expect(icon(3).attributes('data-icon')).toBe('5Quarters:4')
    expect(td(w, 1, 1).find('.xs-in').element.firstElementChild!.tagName.toLowerCase()).toBe('svg')
    expect(td(w, 1, 1).text()).toBe('97')
  })

  it('hides the value where a rule says so: the bar or the icon alone, placed by the alignment', async () => {
    const book: MemoryBook = {
      styles: [{}, { alignment: { horizontal: 'center' } }],
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: 40, appearance: { bar: { start: 0, end: 0.4, color: '#008AEF', gradient: true }, hideValue: true } },
            A2: { value: 50, appearance: { icon: { set: '5Rating', index: 3 }, hideValue: true } },
            A3: { value: 60, style: 1, appearance: { icon: { set: '3Flags', index: 1 }, hideValue: true } },
          },
        },
      ],
    }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    expect(td(w, 1, 1).text()).toBe('')
    expect(css(w, 1, 1).backgroundImage).toContain('#008AEF')
    expect(td(w, 2, 1).text()).toBe('')
    expect(td(w, 2, 1).find('svg.xs-cf-icon').attributes('data-icon')).toBe('5Rating:3')
    expect(getComputedStyle(td(w, 2, 1).find('.xs-in').element).justifyContent).not.toBe('center')
    expect(getComputedStyle(td(w, 3, 1).find('.xs-in').element).justifyContent).toBe('center')
    // the value stays in the tooltip
    expect(td(w, 2, 1).attributes('title')).toBe('50')
  })
})

describe('conditional formats: the local source', () => {
  const book = () => ({
    sheets: [
      {
        name: 'S',
        cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 } },
        conditionalFormats: [{ type: 'colorScale' as const, ranges: ['A1:A3'], points: [{ value: { type: 'min' as const }, color: '#F8696B' }, { value: { type: 'max' as const }, color: '#63BE7B' }] }],
      },
    ],
  })

  it('gives the appearance in its windows', async () => {
    const s = localSource(book(), { calc: 'inline' })
    const got = await s.window({ sheet: 'S', row: 1, rows: 3, col: 1, cols: 1 })
    if ('pending' in got) throw new Error('pending')
    expect(got.appearance).toEqual({ A1: { style: { fill: { color: '#F8696B' } } }, A2: { style: { fill: { color: '#AE9373' } } }, A3: { style: { fill: { color: '#63BE7B' } } } })
    await s.close()
  })

  /** A calculator in the page's thread whose sessions count the rectangles `getCells` reads. */
  function counting(): { calc: Calc; reads: string[] } {
    const inner = createInlineCalc()
    const reads: string[] = []
    const calc: Calc = {
      compute: (b, o) => inner.compute(b, o),
      close: () => inner.close(),
      async createSession(b, o) {
        const session = await inner.createSession(b, o)
        return new Proxy(session, {
          get(target, key) {
            const value = Reflect.get(target, key, target)
            if (key === 'getCells') return (sheet: string, range: string | string[]) => (reads.push([range].flat().join(' ')), (value as AsyncSession['getCells']).call(target, sheet, range))
            return typeof value === 'function' ? value.bind(target) : value
          },
        })
      },
    }
    return { calc, reads }
  }

  it('keeps the values it read for search and the filter after an edit that only repaints, unless a rule formats numbers', async () => {
    const filter = { head: 0, cols: { 1: { text: '1' } } }
    for (const [numberFormat, again] of [[undefined, 0], ['0.00', 1]] as const) {
      const b: BookInput = book()
      if (numberFormat) b.sheets[0]!.conditionalFormats!.push({ type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['5'], style: { numberFormat } })
      const { calc, reads } = counting()
      const s = localSource(b, { calc })
      const events: SourceEvent[] = []
      s.subscribe((e) => events.push(e))
      expect((await s.visibleRows('S', filter)).rows).toEqual([1])
      const before = reads.length
      expect(await s.edit([{ sheet: 'S', cell: 'A3', text: '11' }])).toMatchObject({ ok: true })
      expect(events.map((e) => e.kind)).toEqual(['changed', 'repaint'])
      expect((await s.visibleRows('S', filter)).rows).toEqual([1, 3])
      // the sheet's values read again (A1:A3) only where a rule's format may have changed the shown text
      expect(reads.slice(before).filter((r) => r === 'A1:A3')).toHaveLength(again)
      await s.close()
      await calc.close()
    }
  })

  it('reads the colours again after an edit that recolours another cell', async () => {
    const s = localSource(book(), { calc: 'inline' })
    const w = await mountSheet(s)
    expect(css(w, 2, 1).backgroundColor).toBe('#AE9373')
    await goTo(w, 'A3')
    await key(w, '1')
    await type(w, '11')
    await editorKey(w, 'Enter')
    await settle(12)
    // A2 is now near the scale's low end: its colour changed though the edit did not touch it
    expect(css(w, 2, 1).backgroundColor).toBe('#EA716C')
    await s.close()
  })
})
