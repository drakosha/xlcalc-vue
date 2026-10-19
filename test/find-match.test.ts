// @vitest-environment node
// Excel's Find and Replace over cells' texts, as ./local and memorySource share it: the wildcards, match case, the entire
// cell, a range, the display filter's rows, the order by rows or columns, and Replace All's new texts.
import { describe, expect, it } from 'vitest'
import { findCells, replaceCells, type EachCell } from '../src/find-match'
import type { FindRequest } from '../src/contract/source'

/** Cells by address with their texts, as a source visits them (in no particular order). */
function cells(texts: Record<string, string>): EachCell {
  const list = Object.entries(texts).map(([addr, text]) => {
    const m = /^([A-Z]+)(\d+)$/.exec(addr)!
    const col = [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
    return { row: Number(m[2]), col, text }
  })
  return (visit) => {
    for (const c of [...list].reverse()) visit(c.row, c.col, c.text)
  }
}
const req = (query: string, more: Partial<FindRequest> = {}): FindRequest => ({ sheet: 'S', query, ...more })
/** Whether a query (with options) finds the text. */
const hit = (query: string, text: string, more: Partial<FindRequest> = {}) => findCells(cells({ A1: text }), req(query, more), null, 10).total === 1

describe('the query as Excel reads it', () => {
  it('* stands for any characters, ? for one, inside a cell’s text', () => {
    expect(hit('a*e', 'apple')).toBe(true)
    expect(hit('a*e', 'ae')).toBe(true)
    expect(hit('a*e', 'ap')).toBe(false)
    expect(hit('a?c', 'xabcx')).toBe(true)
    expect(hit('a?c', 'ac')).toBe(false)
    expect(hit('a?c', 'abbc')).toBe(false)
    expect(hit('a*b', 'a\nb')).toBe(true)
    expect(hit('a?b', 'a\nb')).toBe(true)
  })

  it('? is one UTF-16 unit, so a character outside the BMP takes two, as Excel’s Find counts', () => {
    expect(hit('?', '\u{1F600}', { wholeCell: true })).toBe(false)
    expect(hit('??', '\u{1F600}', { wholeCell: true })).toBe(true)
  })

  it('~ before *, ? or ~ is that character; a ~ at the end is dropped, as the library reads a lookup', () => {
    expect(hit('5~*2', '=5*2')).toBe(true)
    expect(hit('5~*2', '502')).toBe(false)
    expect(hit('what~?', 'what?')).toBe(true)
    expect(hit('what~?', 'whats')).toBe(false)
    expect(hit('a~~b', 'a~b')).toBe(true)
    expect(hit('a~~b', 'ab')).toBe(false)
    expect(hit('ab~', 'xab')).toBe(true)
  })

  it('takes every other character as it is, regular expression signs too', () => {
    expect(hit('(1+2)', '=(1+2)*3')).toBe(true)
    expect(hit('(1+2)', '=12')).toBe(false)
    expect(hit('[a]', 'x[a]x')).toBe(true)
    expect(hit('$', 'cost $5')).toBe(true)
    expect(hit('.', 'abc')).toBe(false)
  })

  it('ignores the case without Match case, Cyrillic too; with it the case counts', () => {
    expect(hit('APPLE', 'apple pie')).toBe(true)
    expect(hit('сумм', '=СУММ(A1:A3)')).toBe(true)
    expect(hit('APPLE', 'apple pie', { matchCase: true })).toBe(false)
    expect(hit('apple', 'apple pie', { matchCase: true })).toBe(true)
    expect(hit('ΟΔΟΣ', 'οδος')).toBe(true)
    expect(hit('σ', 'ς', { wholeCell: true })).toBe(true)
  })

  it('with Match entire cell contents the whole text matches the query', () => {
    expect(hit('apple', 'Apple', { wholeCell: true })).toBe(true)
    expect(hit('apple', 'apple pie', { wholeCell: true })).toBe(false)
    expect(hit('apple*', 'apple pie', { wholeCell: true })).toBe(true)
    expect(hit('?', 'a', { wholeCell: true })).toBe(true)
    expect(hit('?', 'ab', { wholeCell: true })).toBe(false)
  })

  it('an empty query, or ~ alone, finds nothing; an empty cell is never a hit', () => {
    expect(findCells(cells({ A1: '~' }), req('~'), null, 10)).toEqual({ hits: [], total: 0 })
    expect(findCells(cells({ A1: 'a' }), req(''), null, 10)).toEqual({ hits: [], total: 0 })
    expect(findCells(cells({ A1: '' }), req('*'), null, 10)).toEqual({ hits: [], total: 0 })
  })
})

describe('where and in which order', () => {
  const texts: Record<string, string> = { '1:1': 'x1', '1:3': 'x2', '2:2': 'x3', '3:1': 'x4', '3:4': 'nothing', '4:2': 'x5' }
  const sheet = cells({ A1: 'x1', C1: 'x2', B2: 'x3', A3: 'x4', D3: 'nothing', B4: 'x5' })
  /** The hits' texts, in their order. */
  const textsOf = (got: { hits: { row: number; col: number }[] }) => got.hits.map((h) => texts[`${h.row}:${h.col}`])

  it('by rows by default, by columns when asked', () => {
    expect(findCells(sheet, req('x'), null, 10)).toEqual({
      hits: [
        { row: 1, col: 1 },
        { row: 1, col: 3 },
        { row: 2, col: 2 },
        { row: 3, col: 1 },
        { row: 4, col: 2 },
      ],
      total: 5,
    })
    expect(textsOf(findCells(sheet, req('x', { byColumns: true }), null, 10))).toEqual(['x1', 'x4', 'x3', 'x5', 'x2'])
  })

  it('only within a range, whole columns and rows too', () => {
    expect(textsOf(findCells(sheet, req('x', { range: 'B1:C3' }), null, 10))).toEqual(['x2', 'x3'])
    expect(textsOf(findCells(sheet, req('x', { range: 'B:B' }), null, 10))).toEqual(['x3', 'x5'])
    expect(textsOf(findCells(sheet, req('x', { range: '3:4' }), null, 10))).toEqual(['x4', 'x5'])
  })

  it('only the rows a display filter shows', () => {
    expect(textsOf(findCells(sheet, req('x'), new Set([1, 4]), 10))).toEqual(['x1', 'x2', 'x5'])
  })

  it('gives at most `max` hits and counts them all', () => {
    expect(findCells(sheet, req('x'), null, 2)).toEqual({ hits: [{ row: 1, col: 1 }, { row: 1, col: 3 }], total: 5 })
  })
})

describe('Replace All’s new texts', () => {
  const replace = (texts: Record<string, string>, query: string, by: string, more: Partial<FindRequest> = {}) =>
    replaceCells(cells(texts), req(query, more), by, null).map((c) => [c.row, c.col, c.text])

  it('replaces every match in a cell with the text as it is, in any case without Match case', () => {
    expect(replace({ A1: 'banana' }, 'a', 'o')).toEqual([[1, 1, 'bonono']])
    expect(replace({ A1: 'Aa' }, 'a', 'b')).toEqual([[1, 1, 'bb']])
    expect(replace({ A1: 'Aa' }, 'a', 'b', { matchCase: true })).toEqual([[1, 1, 'Ab']])
    expect(replace({ A1: 'price' }, 'price', '$& $1')).toEqual([[1, 1, '$& $1']])
    expect(replace({ A1: '=СУММ(A1:A3)' }, 'a3', 'A4')).toEqual([[1, 1, '=СУММ(A1:A4)']])
  })

  it('a * takes as little as it can but the query’s last one, which takes the rest, as Excel’s Replace', () => {
    expect(replace({ A1: 'banana' }, 'a*', 'X')).toEqual([[1, 1, 'bX']])
    expect(replace({ A1: 'abc' }, '*', 'X')).toEqual([[1, 1, 'X']])
    expect(replace({ A1: 'a1b2b' }, 'a*b', 'X')).toEqual([[1, 1, 'X2b']])
  })

  it('with Match entire cell contents the whole text goes', () => {
    expect(replace({ A1: 'apple', A2: 'apple pie' }, 'apple', 'pear', { wholeCell: true })).toEqual([[1, 1, 'pear']])
  })

  it('leaves out cells whose text would not change, and keeps to the range and the shown rows', () => {
    expect(replace({ A1: 'a', B1: 'b' }, 'a', 'a')).toEqual([])
    expect(replaceCells(cells({ A1: 'a', A2: 'a', B2: 'a' }), req('a', { range: 'A1:A2' }), 'z', new Set([2])).map((c) => [c.row, c.col, c.text])).toEqual([[2, 1, 'z']])
  })
})
