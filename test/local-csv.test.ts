// @vitest-environment node
// CSV files in the local source (xlcalc's `xlcalc/csv`): a ru-RU file with `;`, a decimal comma, a date and a UTF-8 byte order
// mark opened into a workbook and calculated, a sheet saved back as Excel saves it, a File naming its sheet.
import { afterEach, describe, expect, it } from 'vitest'
import { openCsv, saveCsv } from '../src/local'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

const BOM = [0xef, 0xbb, 0xbf]
/** A ru-RU file as Excel for Windows saves CSV UTF-8: the mark, `;` between fields, a decimal comma, a date, CRLF. */
const ruFile = (): Uint8Array => new Uint8Array([...BOM, ...new TextEncoder().encode('Товар;Цена;Дата\r\nчай;1234,5;15.03.2026\r\nкофе;"2,25";16.03.2026\r\n')])

describe('openCsv', () => {
  it('reads a ru-RU file with a byte order mark at `;`: numbers with a decimal comma, a date with its format, the workbook in ru-RU', async () => {
    const read = await openCsv(ruFile(), { locale: 'ru-RU', name: 'Цены.csv' })
    expect(read.delimiter).toBe(';')
    expect(read.encoding).toBe('utf-8')
    expect(read.issues).toEqual([])
    expect(read.book.locale).toBe('ru-RU')
    const sheet = read.book.sheets[0]!
    expect(sheet.name).toBe('Цены')
    expect(sheet.cells).toMatchObject({ A1: { value: 'Товар' }, B2: { value: 1234.5 }, B3: { value: 2.25 }, C2: { value: 46096, numberFormat: 14 } })
  })

  it('the workbook opens in a source and calculates; a formula typed over it reads the numbers', async () => {
    const { book } = await openCsv(ruFile(), { locale: 'ru-RU', name: 'Цены.csv' })
    const s = source(book)
    await s.edit([{ sheet: 'Цены', cell: 'D2', text: '=B2+B3' }])
    const w = await win(s, 'Цены', 1, 3, 1, 4)
    expect(w.values[1]![3]).toBe(1236.75)
    expect(w.text[1]![3]).toBe('1236,75')
    expect(w.values[1]![2]).toBe(46096)
  })

  it('a File names the sheet after itself; a Blob and text are read too', async () => {
    const file = new File([ruFile() as BlobPart], 'Март.csv')
    expect((await openCsv(file, { locale: 'ru-RU' })).book.sheets[0]!.name).toBe('Март')
    expect((await openCsv(new Blob([ruFile() as BlobPart]), { locale: 'ru-RU' })).book.sheets[0]!.name).toBe('Sheet1')
    expect((await openCsv('a,b\r\n1.5,2', { name: 'x.csv' })).book.sheets[0]!.cells).toMatchObject({ A2: { value: 1.5 } })
  })

  it('keeps a formula as text unless `formulas` is true, as xlcalc reads a file from elsewhere', async () => {
    expect((await openCsv('=1+2,3')).book.sheets[0]!.cells.A1).toEqual({ value: '=1+2' })
    expect((await openCsv('=1+2,3', { formulas: true })).book.sheets[0]!.cells.A1).toEqual({ formula: '1+2' })
  })

  it('a locale tag without a module fails with the library’s message', async () => {
    await expect(openCsv(ruFile(), { locale: 'de-DE' })).rejects.toThrow(/unknown locale "de-DE"/)
  })
})

describe('saveCsv', () => {
  it("writes a sheet with its formulas' values in the workbook's locale, `;` and the byte order mark, as Excel's CSV UTF-8", async () => {
    const { book } = await openCsv(ruFile(), { locale: 'ru-RU', name: 'Цены.csv' })
    const s = source(book)
    await s.edit([{ sheet: 'Цены', cell: 'D2', text: '=B2+B3' }])
    const csv = await saveCsv(s)
    expect([...csv.data.slice(0, 3)]).toEqual(BOM)
    expect(csv.text).toBe('Товар;Цена;Дата;\r\nчай;1234,5;15.03.2026;1236,75\r\nкофе;2,25;16.03.2026;')
    expect(csv.issues).toEqual([])
  })

  it('writes the sheet asked for and reports the others; in Manual it calculates first', async () => {
    const s = source({ calculation: 'manual', sheets: [{ name: 'A', cells: { A1: { value: 1 } } }, { name: 'B', cells: { A1: { value: 2 }, B1: { formula: 'A1*3' } } }] })
    await s.edit([{ sheet: 'B', cell: 'A1', text: '5' }])
    const csv = await saveCsv(s, { sheet: 'B', bom: false })
    expect(csv.text).toBe('5,15')
    expect(csv.issues.map((i) => i.code)).toEqual(['sheets'])
  })
})
