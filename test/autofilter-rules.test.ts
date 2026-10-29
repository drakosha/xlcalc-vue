// The AutoFilter's rules without the grid: the display filter over its rows (`SheetFilter.last`), and the range Data > Filter
// picks, checked against Excel 16 for Mac (Range.AutoFilter through AppleScript, en-US) on the cases below.
import { describe, expect, it } from 'vitest'
import { cellAddress, parseCell, parseRange, rangeAddress } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import { autoFilterRange } from '../src/core/autofilter-range'
import { filterKey } from '../src/filter-key'
import { columnList, visibleList, type QuerySheet } from '../src/sheet-query'

/** A sheet of texts by row: column A is the row's letter key, column B a status. */
function sheet(rows: string[][]): QuerySheet {
  return {
    maxRow: rows.length,
    textAt: (r, c) => rows[r - 1]?.[c - 1] ?? '',
    valueAt: (r, c) => rows[r - 1]?.[c - 1] || null,
  }
}

describe('the display filter over an AutoFilter (`last`)', () => {
  // a title row, the AutoFilter A2:B6 (header on row 2), and two rows below it
  const s = sheet([['Title'], ['Key', 'Status'], ['k3', 'a'], ['k4', 'b'], ['k5', 'a'], ['k6', 'b'], ['k7', 'b'], ['k8', 'a']])

  it("applies the conditions to the range's rows only; the rows above and below stay in their order", () => {
    expect(visibleList(s, { head: 2, last: 6, cols: { 2: { values: ['a'] } } })).toEqual([1, 2, 3, 5, 7, 8])
  })

  it("sorts the range's rows only", () => {
    expect(visibleList(s, { head: 2, last: 6, sort: { col: 1, dir: 'desc' } })).toEqual([1, 2, 6, 5, 4, 3, 7, 8])
  })

  it("lists the values of the range's rows only", () => {
    expect(columnList(s, 2, { head: 2, last: 6 }).values).toEqual([
      { text: 'a', count: 2 },
      { text: 'b', count: 2 },
    ])
  })

  it('without `last` the conditions reach the end of the sheet, as before', () => {
    expect(visibleList(s, { head: 2, cols: { 2: { values: ['a'] } } })).toEqual([1, 2, 3, 5, 8])
  })

  it('a filter with another last row is another list', () => {
    expect(filterKey({ head: 2, last: 6 })).not.toBe(filterKey({ head: 2, last: 7 }))
  })
})

/** Data > Filter's range for a sheet whose cells with data are `cells` (ranges), from `selection` and the active cell. */
function pick(cells: string[], selection: string): string | null {
  const filled = new Set<string>()
  for (const c of cells) {
    const r = parseRange(c)!
    for (let row = r.top; row <= r.bottom; row++) for (let col = r.left; col <= r.right; col++) filled.add(cellAddress(row, col))
  }
  const box = (a: string[]): CellRect => {
    const at = a.map((x) => parseCell(x)!)
    return { top: Math.min(...at.map((p) => p.row)), left: Math.min(...at.map((p) => p.col)), bottom: Math.max(...at.map((p) => p.row)), right: Math.max(...at.map((p) => p.col)) }
  }
  const data = box([...filled])
  const sel = parseRange(selection)!
  const rect: CellRect = sel.kind === 'cells' ? sel : sel.kind === 'columns' ? { top: 1, bottom: 1_048_576, left: sel.left, right: sel.right } : { top: sel.top, bottom: sel.bottom, left: 1, right: 16_384 }
  const got = autoFilterRange(rect, { top: 1, left: 1, bottom: data.bottom, right: data.right }, (row, col) => filled.has(cellAddress(row, col)))
  return got && rangeAddress(got)
}

describe("Data > Filter's range, as Excel picks it", () => {
  // a header A1:C1 over four rows of data
  const std = ['A1:C5']
  const cases: [name: string, cells: string[], selection: string, excel: string | null][] = [
    ['a cell in the data: its current region', std, 'B3', 'A1:C5'],
    ['an empty cell next to the data: the region without its empty column', std, 'D3', 'A1:C5'],
    ['an empty cell under the data: the region without its empty row', std, 'B6', 'A1:C5'],
    ['an empty cell diagonally off the corner', std, 'D6', 'A1:C5'],
    ['an empty cell inside the data', ['A1:C2', 'A3', 'C3', 'A4:C5'], 'B3', 'A1:C5'],
    ['an empty cell away from the data: none', std, 'F10', null],
    ['a lone cell', ['B2'], 'B2', 'B2'],
    ['a header without data under it', ['B2:D2'], 'C2', 'B2:D2'],
    ['the header row selected: down to the end of the data', std, 'A1:C1', 'A1:C5'],
    ['part of the header row', std, 'A1:B1', 'A1:B5'],
    ['a block inside the data: down to the end of the data', std, 'B2:C3', 'B2:C5'],
    ['a block at the bottom', std, 'B4:C5', 'B4:C5'],
    ['a block larger than the data: trimmed to it', std, 'A1:C9', 'A1:C5'],
    ['a block wider than the data: trimmed, then down', std, 'A1:E2', 'A1:C5'],
    ['an empty block next to the data: none', std, 'D1:D3', null],
    ['an empty block away from the data: none', std, 'E1:F2', null],
    ['an empty block under the data: none', std, 'A6:C7', null],
    ['an entire column', std, 'B:B', 'B1:B5'],
    ['entire columns', std, 'A:B', 'A1:B5'],
    ['an entire row', std, '1:1', 'A1:C5'],
    ['entire rows', std, '2:3', 'A2:C5'],
    ['the entire sheet', std, 'A1:XFD1048576', 'A1:C5'],
    ['a gap of empty rows ends the data', ['A1:C5', 'A8:C9'], 'B3', 'A1:C5'],
    ['the header row above a gap', ['A1:C5', 'A8:C9'], 'A1:C1', 'A1:C5'],
    ['a column further right beyond an empty one is not reached', ['A1:C5', 'E1:E9'], 'A1:C1', 'A1:C5'],
    ['a longer column next to the block takes the region down', ['A1:C5', 'D1:D8'], 'B2:C3', 'B2:C8'],
    ['a longer column next to a cell widens its region', ['A1:C5', 'D1:D8'], 'B3', 'A1:D8'],
    ['a hole in one column does not end the region', ['A1:C3', 'A4', 'C4', 'A5:C5'], 'B2:B3', 'B2:B5'],
    ['an empty row under the block ends it', ['A1:C3', 'A5:C5'], 'B2:C3', 'B2:C3'],
  ]
  it.each(cases)('%s', (_name, cells, selection, excel) => {
    expect(pick(cells, selection)).toBe(excel)
  })
})
