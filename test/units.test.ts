// Pure pieces: addresses, Excel units, prefix sums, the style's CSS and clipboard text.
import { describe, expect, it } from 'vitest'
import { cellAddress, colLabel, colNumber, parseCell, parseRect, rectAddress } from '../src/core/a1'
import { alignCss, borderCss, borderWidth, fillText, fontCss, generalAlign, horizontalOf, verticalOf, wrapsOf } from '../src/core/cell-style'
import { parseTsv, toTsv } from '../src/core/clipboard'
import { numberEditText } from '../src/core/editing'
import { colPx, fontPx, indexAt, prefixSums, rowPx, widthFromPx } from '../src/core/geometry'

describe('addresses', () => {
  it('letters both ways', () => {
    expect([1, 26, 27, 52, 703, 16384].map(colLabel)).toEqual(['A', 'Z', 'AA', 'AZ', 'AAA', 'XFD'])
    expect(['A', 'z', 'AA', 'XFD'].map(colNumber)).toEqual([1, 26, 27, 16384])
    expect(cellAddress(12, 2)).toBe('B12')
  })
  it('cells and rectangles', () => {
    expect(parseCell('$B$12')).toEqual({ row: 12, col: 2 })
    expect(parseCell('B0')).toBeNull()
    expect(parseRect('D4:b2')).toEqual({ r1: 2, c1: 2, r2: 4, c2: 4 })
    expect(parseRect('C3')).toEqual({ r1: 3, c1: 3, r2: 3, c2: 3 })
    expect(parseRect('A1:B2:C3')).toBeNull()
    expect(rectAddress({ r1: 2, c1: 2, r2: 4, c2: 4 })).toBe('B2:D4')
  })
})

describe('Excel units', () => {
  it('column widths as Excel turns them into pixels and back', () => {
    expect(colPx(null)).toBe(64)
    expect(colPx(0)).toBe(0)
    expect(colPx(20.7109375)).toBe(145)
    expect(widthFromPx(64)).toBe(9.140625)
    for (const px of [20, 37, 64, 100, 333]) expect(colPx(widthFromPx(px))).toBe(px)
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
      textDecorationLine: 'underline line-through',
      textDecorationStyle: 'double',
      color: '#111111',
      backgroundColor: '#ffff00',
    })
    expect(fontCss({ font: { color: '#111111' } }, '#ff0000').color).toBe('#ff0000')
  })
  it('alignment and indent of the content box; Fill repeats the text', () => {
    expect(alignCss('right', 2)).toEqual({ textAlign: 'right', paddingRight: '18px' })
    expect(alignCss('left', 1)).toEqual({ textAlign: 'left', paddingLeft: '9px' })
    expect(alignCss('distributed', 0)).toEqual({ textAlign: 'justify', textAlignLast: 'justify' })
    expect(fillText('ab', 40).startsWith('ababab')).toBe(true)
    expect(fillText('', 40)).toBe('')
  })
})

describe('clipboard text', () => {
  it('copies TSV with breaks and tabs inside cells flattened', () => {
    expect(toTsv([['a', 'b\tc'], ['d\ne', '']])).toBe('a\tb c\nd e\t')
  })
  it('reads TSV with quoted fields as Excel writes it', () => {
    expect(parseTsv('a\tb\r\nc\td\r\n')).toEqual([['a', 'b'], ['c', 'd']])
    expect(parseTsv('"x\ny"\t"say ""hi"""\n1\t')).toEqual([['x\ny', 'say "hi"'], ['1', '']])
    expect(parseTsv('5" disk\t2')).toEqual([['5" disk', '2']])
    expect(parseTsv('')).toEqual([])
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
