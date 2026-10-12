// @vitest-environment node
// xlsx in and out of the local source: a book written and read back, a Blob, the issues in a locale and grouped.
import { crc32 } from 'node:zlib'
import { afterEach, describe, expect, it } from 'vitest'
import { compute, sameValue, type CellValue } from 'xlcalc'
import { synthBook } from 'xlcalc/unstable/synth'
import { writeXlsx } from 'xlcalc/xlsx'
import { groupIssues, openXlsx, saveXlsx } from '../src/local'
import { BOLD, closeAll, source, styleOf, win } from './local-fixture'

afterEach(closeAll)

/** A zip of stored entries: enough for a hand-made xlsx with a part the reader reports. */
function storedZip(files: Record<string, string>): Uint8Array {
  const parts: Buffer[] = []
  const dir: Buffer[] = []
  let at = 0
  for (const [name, text] of Object.entries(files)) {
    const n = Buffer.from(name)
    const d = Buffer.from(text)
    const head = Buffer.alloc(30)
    head.writeUInt32LE(0x04034b50, 0)
    head.writeUInt16LE(20, 4)
    head.writeUInt32LE(crc32(d), 14)
    head.writeUInt32LE(d.length, 18)
    head.writeUInt32LE(d.length, 22)
    head.writeUInt16LE(n.length, 26)
    const entry = Buffer.alloc(46)
    entry.writeUInt32LE(0x02014b50, 0)
    entry.writeUInt16LE(20, 4)
    entry.writeUInt16LE(20, 6)
    head.copy(entry, 16, 14, 26)
    entry.writeUInt16LE(n.length, 28)
    entry.writeUInt32LE(at, 42)
    parts.push(head, n, d)
    dir.push(entry, n)
    at += head.length + n.length + d.length
  }
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(dir.length / 2, 8)
  end.writeUInt16LE(dir.length / 2, 10)
  end.writeUInt32LE(dir.reduce((sum, b) => sum + b.length, 0), 12)
  end.writeUInt32LE(at, 16)
  return new Uint8Array(Buffer.concat([...parts, ...dir, end]))
}

const RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
/** A file of one empty sheet `S` with sheet protection, which the book cannot hold. */
const protectedFile = () =>
  storedZip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${RELS}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><sheetProtection sheet="1"/></worksheet>',
  })

describe('xlsx in and out', () => {
  it('writes the book and reads it back with the same values, styles, notes and frozen panes', async () => {
    const s = source()
    await s.edit([{ sheet: 'Лист1', cell: 'A1', text: '1,5' }])
    const { data, issues } = await saveXlsx(s)
    expect(data.byteLength).toBeGreaterThan(1000)
    expect(issues).toEqual([])
    const back = await openXlsx(data)
    expect(back.book.locale).toBe('ru-RU')
    const t = source(back.book)
    const w = await win(t, 'Лист1', 1, 5, 1, 7)
    expect(w.text.map((r) => r[0])).toEqual(['1,5', '2', '3,5', '', ''])
    expect(w.formulas.A3).toBe('СУММ(A1:A2)')
    expect(styleOf(w, 'C1')).toMatchObject(BOLD)
    expect(w.notes?.B1).toEqual({ author: 'Аня', text: 'Проверить' })
    expect((await t.sheets())[0]!.freeze).toEqual({ rows: 1, cols: 1 })
    expect((await win(t, 'Таблица', 2, 1, 3, 1)).formulas.C2).toBe('[@Сумма]*2')
  })

  it('keeps every value through a file: opened, edited, saved, the file\'s stored values are the session\'s', async () => {
    const { book } = synthBook({ seed: 3, rows: 40 })
    const first = await openXlsx((await writeXlsx({ book, result: compute(book) })).data)
    const s = source(first.book)
    expect(await s.edit([{ sheet: 'Данные', cell: 'D2', text: '123456.78' }, { sheet: 'Данные', cell: 'D3', text: '=D2/2' }])).toMatchObject({ ok: true })
    expect(await s.paste('Данные', 'D4', '10\r\n20')).toMatchObject({ ok: true })
    const before = await s.snapshot()
    const back = await openXlsx((await saveXlsx(s)).data)
    // a formula's or a spill's value as the file stores it (cached, nothing computed), exactly; a constant as the book
    // read holds it, which keeps the full double where the result gives 15 significant digits (sameValue)
    const mismatches: [string, string, unknown, CellValue][] = []
    let cached = 0
    let constants = 0
    before.result.sheets.forEach((sheet, i) => {
      for (const [addr, cell] of Object.entries(sheet.cells)) {
        if (cell.value === null) continue
        const stored = back.cached.sheets[i]?.cells[addr]
        if (stored) cached++
        else constants++
        const same = stored ? JSON.stringify(stored.value) === JSON.stringify(cell.value) : sameValue(back.book.sheets[i]?.cells[addr]?.value as CellValue, cell.value)
        if (!same) mismatches.push([sheet.name, addr, stored ? stored.value : back.book.sheets[i]?.cells[addr]?.value, cell.value])
      }
    })
    expect(mismatches).toEqual([])
    expect(cached).toBeGreaterThan(2000)
    expect(constants).toBeGreaterThan(500)
  })

  it('opens a circular reference holding the values its file stores, with the session\'s `cached`, as Excel opens the file', async () => {
    const cells = { A1: { formula: 'B1/2+1' }, B1: { formula: 'A1' }, C1: { formula: 'A1*10' }, D1: { value: 5 }, E1: { formula: 'D1*2' } }
    const stored = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { value: 2 }, C1: { value: 20 }, E1: { value: 10 } } }] }
    const read = await openXlsx((await writeXlsx({ book: { sheets: [{ name: 'S', cells }] }, result: stored })).data)
    // without them the cycle and its reader hold 0; a formula outside it is computed either way
    expect((await win(source(read.book), 'S', 1, 1, 1, 5)).values).toEqual([[0, 0, 0, 5, 10]])
    const s = source(read.book, { session: { cached: read.cached } })
    expect((await win(s, 'S', 1, 1, 1, 5)).values).toEqual([[2, 2, 20, 5, 10]])
    // an edit of another cell keeps them, and the saved file stores them again
    expect(await s.edit([{ sheet: 'S', cell: 'D1', text: '7' }])).toMatchObject({ ok: true })
    expect((await win(s, 'S', 1, 1, 1, 5)).values).toEqual([[2, 2, 20, 7, 14]])
    const back = await openXlsx((await saveXlsx(s)).data)
    expect(back.cached.sheets[0]!.cells).toMatchObject({ A1: { value: 2 }, B1: { value: 2 }, C1: { value: 20 }, E1: { value: 14 } })
  })

  it('opens a Blob as a file input gives it', async () => {
    const { data } = await saveXlsx(source())
    const back = await openXlsx(new Blob([data as BlobPart]))
    expect(back.book.sheets.map((x) => x.name)).toEqual(['Лист1', 'Таблица', 'Список'])
    expect(back.issues).toEqual([])
  })

  it('gives the reader\'s and the writer\'s issues in the language of the locale asked for, en-US without one', async () => {
    const file = protectedFile()
    expect((await openXlsx(file)).issues).toEqual([{ code: 'protection', sheet: 'S', message: expect.stringMatching(/^sheet protection/) }])
    expect((await openXlsx(file, { locale: 'ru-RU' })).issues).toEqual([{ code: 'protection', sheet: 'S', message: expect.stringMatching(/^Защита листа/) }])
    const s = source({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 1, meta: { id: 7 } } } }] })
    expect((await saveXlsx(s)).issues).toEqual([{ code: 'meta', sheet: 'S', cells: ['A1'], message: expect.stringMatching(/^cell meta/) }])
    expect((await saveXlsx(s, { locale: 'ru-RU', mode: 'values' })).issues).toEqual([{ code: 'meta', sheet: 'S', cells: ['A1'], message: expect.stringMatching(/^Поле meta/) }])
  })

  it('fails a locale xlcalc has no module for, with the library message', async () => {
    await expect(openXlsx(protectedFile(), { locale: 'de-DE' })).rejects.toThrow(/unknown locale "de-DE"/)
  })

  it('groups what the reader left out by kind, for a notice', () => {
    expect(
      groupIssues([
        { code: 'drawing', message: 'a chart', sheet: 'A' },
        { code: 'conditionalFormat', message: 'a rule', sheet: 'A', cells: ['B2:B9'] },
        { code: 'drawing', message: 'a picture', sheet: 'B' },
        { code: 'drawing', message: 'a chart', sheet: 'A' },
      ]),
    ).toEqual([
      { code: 'drawing', count: 3, sheets: ['A', 'B'], message: 'a chart' },
      { code: 'conditionalFormat', count: 1, sheets: ['A'], message: 'a rule' },
    ])
  })
})
