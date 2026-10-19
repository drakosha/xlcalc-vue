// The example app's opening of a file: the workbook's language when the file specifies none, the sheet shown first, and
// what the file lost, grouped by kind for the notice.
import { describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { XlsxIssue } from 'xlcalc/xlsx'
import { bookLocaleOf, firstShownSheet, lossesOf } from '../example/spa-open'

describe('bookLocaleOf: the language a book opens in', () => {
  it('takes the interface language for a file without one, as Excel shows formulas in its user language', () => {
    expect(bookLocaleOf(undefined, 'ru')).toBe('ru-RU')
    expect(bookLocaleOf(undefined, 'en')).toBe('en-US')
  })
  it('keeps the language a file names (xlcalc writes it), whatever the interface', () => {
    expect(bookLocaleOf('en-US', 'ru')).toBe('en-US')
    expect(bookLocaleOf('ru-RU', 'en')).toBe('ru-RU')
  })
  it('takes the interface language for a language the app has no book locale for', () => {
    expect(bookLocaleOf('de-DE', 'ru')).toBe('ru-RU')
  })
})

describe('firstShownSheet: the sheet the app opens', () => {
  const sheet = (name: string, cells: BookInput['sheets'][number]['cells'], more: object = {}) => ({ name, cells, ...more })
  const book: BookInput = {
    sheets: [
      sheet('Pres->', { A1: { value: 'x' } }, { hidden: true }),
      sheet('Outputs->', {}),
      sheet('Production', { A1: { value: 'Plan' } }),
      sheet('Later', { A1: { value: 1 } }),
    ],
  }
  it('opens the sheet the file saved as active, as Excel does', () => {
    expect(firstShownSheet({ ...book, activeSheet: 'Later' })).toBe('Later')
  })
  it('without a saved active sheet opens the first visible one, an empty divider too, as Excel does', () => {
    expect(firstShownSheet(book)).toBe('Outputs->')
  })
  it('gives nothing for a book without sheets, so the sheet opens as it would', () => {
    expect(firstShownSheet({ sheets: [] })).toBeUndefined()
  })
})

describe('lossesOf: what a file lost, by kind', () => {
  const issues: XlsxIssue[] = [
    { code: 'name', message: 'a hidden name is read as a visible one', name: 'n1' },
    { code: 'name', message: 'a hidden name is read as a visible one', name: 'n2' },
    { code: 'style', message: 'a pattern fill', sheet: 'Data', cells: ['B2', 'C2:D3'] },
    { code: 'externalLink', message: 'links to other workbooks are not in the model' },
    { code: 'externalLink', message: 'references to other workbooks', sheet: 'Лист 1', cells: ['A1', 'A2', 'A3'] },
    { code: 'externalLink', message: 'references to other workbooks', sheet: 'Data', cells: ['E5', 'E6', 'E7', 'E8'] },
    { code: 'externalLink', message: 'references to other workbooks', name: 'Rate' },
    { code: 'name', message: 'a built-in name is not in the model', name: '_xlnm.Print_Area' },
  ]
  const kinds = lossesOf(issues)

  it('puts what changes values first and names last, whatever order the reader met them in', () => {
    expect(kinds.map((k) => k.code)).toEqual(['externalLink', 'style', 'name'])
  })
  it('counts the issues of a kind and each message once with its count, in the order met', () => {
    const link = kinds[0]!
    expect(link.count).toBe(4)
    expect(link.messages).toEqual([
      { text: 'links to other workbooks are not in the model', count: 1 },
      { text: 'references to other workbooks', count: 3 },
    ])
    expect(kinds[2]!.messages).toEqual([
      { text: 'a hidden name is read as a visible one', count: 2 },
      { text: 'a built-in name is not in the model', count: 1 },
    ])
  })
  it('lists the sheets once and the cells with their sheet, quoted as Excel quotes it, the first five as examples', () => {
    const link = kinds[0]!
    expect(link.sheets).toEqual(['Лист 1', 'Data'])
    expect(link.cells).toEqual({ count: 7, examples: ["'Лист 1'!A1", "'Лист 1'!A2", "'Лист 1'!A3", 'Data!E5', 'Data!E6'] })
    expect(link.names).toEqual({ count: 1, examples: ['Rate'] })
    expect(kinds[1]!.cells).toEqual({ count: 2, examples: ['Data!B2', 'Data!C2:D3'] })
    expect(kinds[2]!.names).toEqual({ count: 3, examples: ['n1', 'n2', '_xlnm.Print_Area'] })
  })
  it('counts a cell or a name two issues of a kind are about once', () => {
    const notes = lossesOf([
      { code: 'comment', message: "a note's formatted text is read as its plain text", sheet: 'Data', cells: ['E5'] },
      { code: 'comment', message: 'a note shown all the time is read as one shown on hover', sheet: 'Data', cells: ['E5', 'E6'] },
      { code: 'name', message: 'a hidden name', name: 'Rate' },
      { code: 'name', message: 'a name of a sheet left out', name: 'Rate' },
    ])
    expect(notes.map((k) => [k.cells, k.names])).toEqual([
      [{ count: 2, examples: ['Data!E5', 'Data!E6'] }, { count: 0, examples: [] }],
      [{ count: 0, examples: [] }, { count: 1, examples: ['Rate'] }],
    ])
  })
})
