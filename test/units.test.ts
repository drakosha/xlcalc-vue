// Pure pieces: Excel units, prefix sums and the style's CSS.
import { describe, expect, it } from 'vitest'
import { ref, shallowRef } from 'vue'
import { alignCss, borderCss, borderWidth, decorationCss, fillText, fontCss, generalAlign, horizontalOf, verticalOf, wrapsOf } from '../src/core/cell-style'
import { numberEditText, textEditText } from '../src/core/edit-text'
import { colPx, fontPx, indexAt, prefixSums, rowPx, useGeometry, widthFromPx } from '../src/core/geometry'

describe('Excel units', () => {
  it("column widths in the library's geometry (Excel for Mac) to whole pixels and back", () => {
    // A column without a width of its own is 53 points, 70.67 px; xlcalc lays out `####` and overflow by these.
    expect(colPx(null)).toBe(71)
    expect(colPx(8.83203125)).toBe(71)
    expect(colPx(9.140625)).toBe(73)
    expect(colPx(15)).toBe(120)
    expect(colPx(0)).toBe(0)
    expect(widthFromPx(71)).toBe(8.83203125)
    expect(widthFromPx(120)).toBe(15)
    // a width goes back in whole points, as the library stores it; pixels of whole points come back as they were
    for (const px of [20, 40, 72, 120, 332]) expect(colPx(widthFromPx(px))).toBe(px)
    expect(widthFromPx(4000)).toBe(255)
  })
  it('row heights and font sizes', () => {
    expect(rowPx(15)).toBe(20)
    expect(rowPx(0)).toBe(0)
    expect(fontPx(11)).toBe(12.5)
    expect(fontPx(22)).toBe(25)
  })
  it('prefix sums and the item under an offset', () => {
    const sizes = [0, 10, 0, 30, 20]
    const off = prefixSums(4, (i) => sizes[i]!)
    expect([...off]).toEqual([0, 10, 10, 40, 60])
    expect(indexAt(off, 4, 0)).toBe(1)
    expect(indexAt(off, 4, 9.9)).toBe(1)
    expect(indexAt(off, 4, 10)).toBe(3)
    expect(indexAt(off, 4, 45)).toBe(4)
    expect(indexAt(off, 4, 1e9)).toBe(4)
    expect(indexAt(prefixSums(0, () => 1), 0, 5)).toBe(1)
  })
})

describe('style to CSS', () => {
  it('General aligns by the value as Excel does', () => {
    expect(generalAlign(1)).toBe('right')
    expect(generalAlign('x')).toBe('left')
    expect(generalAlign(true)).toBe('center')
    expect(generalAlign({ error: 'NA' })).toBe('center')
    expect(generalAlign(null)).toBe('left')
    expect(horizontalOf({ alignment: { horizontal: 'general' } }, 5)).toBe('right')
    expect(horizontalOf({ alignment: { horizontal: 'left' } }, 5)).toBe('left')
  })
  it('vertical: the bottom by default; justify and distributed wrap', () => {
    expect(verticalOf(undefined)).toBe('bottom')
    expect(verticalOf({ alignment: { vertical: 'center' } })).toBe('center')
    expect(verticalOf({ alignment: { vertical: 'justify' } })).toBe('top')
    expect(wrapsOf({ alignment: { horizontal: 'justify' } })).toBe(true)
    expect(wrapsOf({ alignment: { horizontal: 'left' } })).toBe(false)
  })
  it('every xlsx border kind has a CSS line', () => {
    const line = (l: Parameters<typeof borderCss>[0]) => Object.values(borderCss(l)!).join(' ')
    expect(line({ style: 'thin' })).toBe('1px solid var(--xs-border-auto)')
    expect(line({ style: 'medium', color: '#ff0000' })).toBe('2px solid #ff0000')
    expect(line({ style: 'thick' })).toBe('3px solid var(--xs-border-auto)')
    expect(line({ style: 'double' })).toBe('3px double var(--xs-border-auto)')
    expect(line({ style: 'hair' })).toBe('1px dotted var(--xs-border-auto)')
    expect(line({ style: 'dotted' })).toBe('1px dotted var(--xs-border-auto)')
    expect(line({ style: 'dashed' })).toBe('1px dashed var(--xs-border-auto)')
    expect(line({ style: 'mediumDashDot' })).toBe('2px dashed var(--xs-border-auto)')
    expect(borderWidth({ style: 'slantDashDot' })).toBe(2)
    expect(borderCss(undefined)).toBeUndefined()
  })
  it('font, colours and fill; the format section colour wins over the font colour', () => {
    expect(fontCss({ font: { name: 'Arial', size: 22, bold: true, italic: true, underline: 'double', strike: true, color: '#111111' }, fill: { color: '#ffff00' } }, undefined)).toEqual({
      fontFamily: '"Arial", var(--xs-font)',
      fontSize: '25px',
      fontWeight: '700',
      fontStyle: 'italic',
      color: '#111111',
      backgroundColor: '#ffff00',
    })
    expect(fontCss({ font: { color: '#111111' } }, '#ff0000').color).toBe('#ff0000')
    expect(decorationCss({ font: { underline: 'double', strike: true } }, false)).toEqual({ textDecorationLine: 'underline line-through', textDecorationStyle: 'double' })
    expect(decorationCss({ font: { strike: true } }, true)).toEqual({ textDecorationLine: 'underline line-through' })
    expect(decorationCss(undefined, false)).toEqual({})
  })
  it('alignment and indent of the content box; Fill repeats the text', () => {
    expect(alignCss('right', 2)).toEqual({ textAlign: 'right', paddingRight: '18px' })
    expect(alignCss('left', 1)).toEqual({ textAlign: 'left', paddingLeft: '9px' })
    expect(alignCss('distributed', 0)).toEqual({ textAlign: 'justify', textAlignLast: 'justify' })
    expect(fillText('ab', 40).startsWith('ababab')).toBe(true)
    expect(fillText('', 40)).toBe('')
  })
})

describe("a number's edit text in the book's locale", () => {
  it('the full value, a percent as a percent, a date as shown', () => {
    expect(numberEditText(1234.5678, '1 234,57 ₽', 'ru-RU')).toBe('1234,5678')
    expect(numberEditText(1234.5678, '1,234.57', 'en-US')).toBe('1234.5678')
    expect(numberEditText(0.355, '35,5%', 'ru-RU')).toBe('35,5%')
    expect(numberEditText(46096, '15.03.2026', 'ru-RU')).toBe('15.03.2026')
    expect(numberEditText(0.75, '18:00', 'ru-RU')).toBe('18:00')
    expect(numberEditText(-42, '(42.00)', 'en-US')).toBe('-42')
  })
})

describe("a text's edit text: an apostrophe where the book would read it as something else", () => {
  // Either locale's forms count in both: an extra apostrophe is harmless, a lost text is not.
  it('numbers, percents, money, exponents in either locale', () => {
    for (const t of ['007', '1,5', '1.5', '-3', '+3', '1 234,5', '1\u00a0234,5', '1,234.5', '(42)', '5%', '12,5 %', '1E+05', '1,2e-3', '$12', '12 ₽', '€3.50'])
      expect([t, textEditText(t)]).toEqual([t, `'${t}`])
    expect(textEditText('1,234.50')).toBe("'1,234.50")
  })
  it('dates and times as either locale writes them', () => {
    for (const t of ['15.03.2026', '15.03.26', '15.03', '3/15/2026', '3/15', '2026-03-15', '1-2', '15 мар 2026', '15 марта', 'Mar 15', '15-Mar-2026', '12:30', '9:05:07', '7:30 PM', '15.03.2026 12:30'])
      expect([t, textEditText(t)]).toEqual([t, `'${t}`])
  })
  it('booleans and errors of both locales, formulas and a leading apostrophe', () => {
    for (const t of ['ИСТИНА', 'ложь', 'TRUE', 'false', '#Н/Д', '#ДЕЛ/0!', '#N/A', '#VALUE!', '=SUM(A1)', '+7 999', '-x', '@name', "'quoted"])
      expect([t, textEditText(t)]).toEqual([t, `'${t}`])
  })
  it('ordinary text stays as it is', () => {
    for (const t of ['007 agent', 'Отдел 5', '5 apples', 'mart', 'Счёт № 12', 'a-1', 'ИСТИНА или ложь', '', '-', 'K-000300', 'v1.2.3'])
      expect([t, textEditText(t)]).toEqual([t, t])
  })
})


describe("rows' offsets", () => {
  /** Geometry over a state of its own: `visible` a display order of sheet rows, or null. */
  function rig(count: number) {
    const heights = shallowRef(new Map<number, number>())
    const visible = shallowRef<number[] | null>(null)
    const rowCount = ref(count)
    const geo = useGeometry({
      rowCount,
      colCount: ref(1),
      bookWidths: shallowRef(new Map()),
      userWidths: shallowRef({}),
      heights,
      defaultRowPx: ref(20),
      visible,
      rowOf: (d) => visible.value?.[d - 1] ?? d,
      indexOf: (r) => (visible.value ? (visible.value.indexOf(r) + 1 || undefined) : r),
    })
    const full = () => prefixSums(rowCount.value, (d) => heights.value.get(visible.value?.[d - 1] ?? d) ?? 20)
    return { geo, heights, visible, rowCount, full }
  }

  it('changed heights sum again from the first changed row, the same as summing all; a filter, a reset and a new count too', () => {
    const { geo, heights, visible, rowCount, full } = rig(500)
    expect([...geo.rowOffsets.value]).toEqual([...full()])
    const steps: (() => void)[] = [
      () => geo.setHeights([[10, 40], [300, 0]]),
      () => geo.setHeights([[10, undefined], [499, 7]]),
      () => geo.setHeights([[1, 30]]),
      () => geo.setHeights([[500, 11], [250, 25], [251, 26]]),
      () => (visible.value = Array.from({ length: 200 }, (_, i) => 500 - i * 2)),
      () => geo.setHeights([[498, 50], [3, 60]]),
      () => (visible.value = null),
      () => (rowCount.value = 520),
      () => geo.setHeights([[515, 33]]),
      () => (heights.value = new Map([[2, 45]])),
      () => geo.setHeights([[2, 45]]),
    ]
    for (const step of steps) {
      step()
      rowCount.value = visible.value ? visible.value.length : Math.max(rowCount.value, 500)
      expect([...geo.rowOffsets.value]).toEqual([...full()])
    }
  })
})
