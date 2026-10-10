// The the app's formula-refs tests, translated, and ru-RU cases: `;` between arguments, a decimal comma, Russian names.
import { describe, expect, it } from 'vitest'
import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, formulaTokens, parseCellRef, sheetRefName, type FormulaRef } from '../src/core/formula-refs'

const refs = (f: string): FormulaRef[] => formulaTokens(f).flatMap((t) => (t.ref ? [t.ref] : []))
const same = (r1: number, c1: number, r2 = r1, c2 = c1): FormulaRef => ({ sheet: null, r1, c1, r2, c2 })

describe('formulaTokens', () => {
  it('single cells and a range', () => {
    expect(refs('B1+B2*2')).toEqual([same(1, 2), same(2, 2)])
    expect(refs('SUM(B1:B4)')).toEqual([same(1, 2, 4, 2)])
    expect(formulaTokens('SUM(B1:B4)').map((t) => t.text)).toEqual(['SUM(', 'B1:B4', ')'])
  })

  it('absolute references and a range given from its bottom-right corner', () => {
    expect(refs('SUM($B$1:$B$2)')).toEqual([same(1, 2, 2, 2)])
    expect(refs('D4:B2')).toEqual([same(2, 2, 4, 4)])
  })

  it('a sheet bare, quoted, with an apostrophe inside', () => {
    expect(refs("Высоты!B3+'Второй лист'!A1")).toEqual([
      { sheet: 'Высоты', r1: 3, c1: 2, r2: 3, c2: 2 },
      { sheet: 'Второй лист', r1: 1, c1: 1, r2: 1, c2: 1 },
    ])
    expect(refs("'Иван''s'!A1")[0]?.sheet).toBe("Иван's")
  })

  it('whole columns and rows', () => {
    expect(refs('SUM(B:B)+MAX(1:1)')).toEqual([
      { sheet: null, r1: 1, c1: 2, r2: EXCEL_MAX_ROWS, c2: 2 },
      { sheet: null, r1: 1, c1: 1, r2: 1, c2: EXCEL_MAX_COLS },
    ])
  })

  it('a function name with digits and range names are not references', () => {
    expect(refs('LOG10(A1)')).toEqual([same(1, 1)])
    expect(refs('ATAN2(1,2)')).toEqual([])
    expect(refs('Итого2024+A1')).toEqual([same(1, 1)])
    expect(refs('_xlfn.XLOOKUP(A1,B:B,C:C)').length).toBe(3)
  })

  it('quoted text is skipped', () => {
    expect(refs('"см. A1"&B2')).toEqual([same(2, 2)])
    expect(refs('"кавычка "" и A1"')).toEqual([])
    expect(formulaTokens('"см. A1"&B2').map((t) => t.text)).toEqual(['"см. A1"&', 'B2'])
  })

  it('a malformed address stays text', () => {
    expect(refs('XFE1')).toEqual([])
    expect(refs('A0')).toEqual([])
  })
})

describe('formulaTokens in ru-RU', () => {
  it('`;` separates arguments and a decimal comma is no reference', () => {
    expect(refs('СУММ(A1:A3;1,5)')).toEqual([same(1, 1, 3, 1)])
    expect(formulaTokens('СУММ(A1:A3;1,5)').map((t) => t.text)).toEqual(['СУММ(', 'A1:A3', ';1,5)'])
    expect(refs('ЕСЛИ(B2>0,5;C2;D2)')).toEqual([same(2, 2), same(2, 3), same(2, 4)])
  })

  it('Russian function names, with digits too, are not references; addresses are not translated', () => {
    expect(refs('ОКРУГЛ(A2/A3*1,5;2)')).toEqual([same(2, 1), same(3, 1)])
    expect(refs('ЛОГ10(B7)')).toEqual([same(7, 2)])
    expect(refs('СУММ(A:A)-СРЗНАЧ(2:2)')).toEqual([
      { sheet: null, r1: 1, c1: 1, r2: EXCEL_MAX_ROWS, c2: 1 },
      { sheet: null, r1: 2, c1: 1, r2: 2, c2: EXCEL_MAX_COLS },
    ])
  })

  it('Russian sheet names, bare and quoted, with `;` after them', () => {
    expect(refs("Лист1!A1;'Длинный лист'!B2:C3")).toEqual([
      { sheet: 'Лист1', r1: 1, c1: 1, r2: 1, c2: 1 },
      { sheet: 'Длинный лист', r1: 2, c1: 2, r2: 3, c2: 3 },
    ])
  })

  it('the text joins back into the formula', () => {
    const f = 'ЕСЛИ(СУММ($A$1:A3;0,5)>1;"да; A1";Лист2!B$4)'
    expect(formulaTokens(f).map((t) => t.text).join('')).toBe(f)
    expect(refs(f).map((r) => [r.sheet, r.r1, r.c1])).toEqual([
      [null, 1, 1],
      ['Лист2', 4, 2],
    ])
  })
})

describe('parseCellRef', () => {
  it('a cell, a range, a sheet with an address, and a sheet alone', () => {
    expect(parseCellRef('B12')).toEqual({ sheet: null, range: { sheet: null, r1: 12, c1: 2, r2: 12, c2: 2 } })
    expect(parseCellRef('b2:D15')?.range).toEqual({ sheet: null, r1: 2, c1: 2, r2: 15, c2: 4 })
    expect(parseCellRef("'Лист 2'!C7")).toEqual({ sheet: 'Лист 2', range: { sheet: 'Лист 2', r1: 7, c1: 3, r2: 7, c2: 3 } })
    expect(parseCellRef('Реестр')).toEqual({ sheet: 'Реестр', range: null })
    expect(parseCellRef("'Второй лист'")).toEqual({ sheet: 'Второй лист', range: null })
    expect(parseCellRef('B12+C1')).toBeNull()
    expect(parseCellRef('')).toBeNull()
  })
  it('a sheet name is quoted only with spaces and signs', () => {
    expect(sheetRefName('Реестр')).toBe('Реестр')
    expect(sheetRefName('Лист 2')).toBe("'Лист 2'")
    expect(sheetRefName("Иван's")).toBe("'Иван''s'")
  })
})
