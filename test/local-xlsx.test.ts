// @vitest-environment node
// xlsx reading and writing in the local source: a workbook written and read back, a Blob, the issues in a locale and grouped.
import { crc32 } from 'node:zlib'
import { afterEach, describe, expect, it } from 'vitest'
import { compute, sameValue, type CellValue } from 'xlcalc'
import { synthBook } from 'xlcalc/unstable/synth'
import { writeXlsx } from 'xlcalc/xlsx'
import { groupIssues, openXlsx, saveXlsx } from '../src/local'
import { BOLD, closeAll, source, styleOf, win } from './local-fixture'
import { unzipText } from './zip'

afterEach(closeAll)

/** A zip of stored entries: enough for a handmade xlsx with a part the reader reports. */
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
/** A file with one empty sheet `S` protected with a password, which the workbook does not hold. */
const protectedFile = () =>
  storedZip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${RELS}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><sheetProtection sheet="1" password="CC3D"/></worksheet>',
  })

/** A file with 1 in `S!A1` and a page break before row 6, which the workbook does not hold. */
const brokenFile = () =>
  storedZip({
    ...unzipText(protectedFile()),
    'xl/worksheets/sheet1.xml':
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks></worksheet>',
  })

describe('xlsx reading and writing', () => {
  it('with keep writes the workbook over the file it was read from, keeping what the workbook does not hold', async () => {
    const file = brokenFile()
    const read = await openXlsx(file, { keep: true })
    const s = source(read.book)
    expect((await s.edit([{ sheet: 'S', cell: 'A1', text: '5' }])).ok).toBe(true)
    const kept = unzipText((await saveXlsx(s, { over: file })).data)['xl/worksheets/sheet1.xml']!
    expect(kept).toContain('<brk id="5"')
    expect(kept).toMatch(/<c r="A1"[^>]*><v>5<\/v>/)
    // Anew, the workbook's own parts only
    expect(unzipText((await saveXlsx(s)).data)['xl/worksheets/sheet1.xml']).not.toContain('<brk')
  })

  it('writes the workbook and reads it back with the same values, styles, notes and frozen panes', async () => {
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

  it("keeps every value through a file: opened, edited and saved, the file's stored values are the session's", async () => {
    const { book } = synthBook({ seed: 3, rows: 40 })
    const first = await openXlsx((await writeXlsx({ book, result: compute(book) })).data)
    const s = source(first.book)
    expect(await s.edit([{ sheet: 'Данные', cell: 'D2', text: '123456.78' }, { sheet: 'Данные', cell: 'D3', text: '=D2/2' }])).toMatchObject({ ok: true })
    expect(await s.paste('Данные', 'D4', '10\r\n20')).toMatchObject({ ok: true })
    const before = await s.snapshot()
    const back = await openXlsx((await saveXlsx(s)).data)
    // a formula's or a spill's value exactly as the file stores it (cached, nothing calculated); a constant as the workbook
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

  it("opens a circular reference with the values its file stores kept, through the session's `cached`, as Excel opens the file", async () => {
    const cells = { A1: { formula: 'B1/2+1' }, B1: { formula: 'A1' }, C1: { formula: 'A1*10' }, D1: { value: 5 }, E1: { formula: 'D1*2' } }
    const stored = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { value: 2 }, C1: { value: 20 }, E1: { value: 10 } } }] }
    const read = await openXlsx((await writeXlsx({ book: { sheets: [{ name: 'S', cells }] }, result: stored })).data)
    // without them the cycle and the cell that refers to it contain 0; a formula outside it is calculated either way
    expect((await win(source(read.book), 'S', 1, 1, 1, 5)).values).toEqual([[0, 0, 0, 5, 10]])
    const s = source(read.book, { session: { cached: read.cached } })
    expect((await win(s, 'S', 1, 1, 1, 5)).values).toEqual([[2, 2, 20, 5, 10]])
    // an edit of another cell keeps them, and the saved file stores them again
    expect(await s.edit([{ sheet: 'S', cell: 'D1', text: '7' }])).toMatchObject({ ok: true })
    expect((await win(s, 'S', 1, 1, 1, 5)).values).toEqual([[2, 2, 20, 7, 14]])
    const back = await openXlsx((await saveXlsx(s)).data)
    expect(back.cached.sheets[0]!.cells).toMatchObject({ A1: { value: 2 }, B1: { value: 2 }, C1: { value: 20 }, E1: { value: 14 } })
  })

  it('opens a Blob as a file input provides it', async () => {
    const { data } = await saveXlsx(source())
    const back = await openXlsx(new Blob([data as BlobPart]))
    expect(back.book.sheets.map((x) => x.name)).toEqual(['Лист1', 'Таблица', 'Список'])
    expect(back.issues).toEqual([])
  })

  it("returns the reader's and the writer's issues in the language of the requested locale, en-US without one", async () => {
    const file = protectedFile()
    expect((await openXlsx(file)).issues).toEqual([{ code: 'protection', sheet: 'S', message: expect.stringMatching(/^the sheet protection's password/) }])
    expect((await openXlsx(file, { locale: 'ru-RU' })).issues).toEqual([{ code: 'protection', sheet: 'S', message: expect.stringMatching(/^Пароль защиты листа/) }])
    const s = source({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 1, meta: { id: 7 } } } }] })
    expect((await saveXlsx(s)).issues).toEqual([{ code: 'meta', sheet: 'S', cells: ['A1'], message: expect.stringMatching(/^a cell's meta/) }])
    expect((await saveXlsx(s, { locale: 'ru-RU', mode: 'values' })).issues).toEqual([{ code: 'meta', sheet: 'S', cells: ['A1'], message: expect.stringMatching(/^Поле meta/) }])
  })

  it("fails for a locale xlcalc has no module for, with the library's message", async () => {
    await expect(openXlsx(protectedFile(), { locale: 'de-DE' })).rejects.toThrow(/unknown locale "de-DE"/)
  })

  it('groups what the reader omitted by kind, for a notice', () => {
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

/** A file `writeXlsx` writes of sheet S with a bold #4472C4 style, its parts by name, to patch before `storedZip`. */
async function partsOfBook(cells: Record<string, { value: string }>): Promise<Record<string, string>> {
  const s = source({ styles: [{ font: { bold: true, color: '#4472C4' } }], sheets: [{ name: 'S', cells: { ...cells, Z99: { value: 0, style: 0 } } as never }] })
  return unzipText((await saveXlsx(s)).data)
}
/** The cell format a written file gives a cell of sheet S: its `<xf>` in the styles. */
function xfOf(files: Record<string, string>, cell: string): string | undefined {
  const xfs = /<cellXfs[^>]*>(.*?)<\/cellXfs>/.exec(files['xl/styles.xml']!)![1]!.match(/<xf\b[^>]*?(?:\/>|>.*?<\/xf>)/g)!
  const tag = new RegExp(`<c r="${cell}"[^>]*>`).exec(files['xl/worksheets/sheet1.xml']!)?.[0]
  return tag && xfs[Number(/ s="(\d+)"/.exec(tag)?.[1] ?? 0)]
}

describe('xlsx over its original file: xlcalc 0.61 to 0.64', () => {
  it("keeps a quote prefix and a theme color of cells inserted rows moved, where the workbook sees the look of the file's first format (0.64.0)", async () => {
    const files = await partsOfBook({ A1: { value: '007' }, A2: { value: 'blue' } })
    // two formats the workbook reads as 0 and 1: plain with a quote prefix, bold in the theme's first accent (#4472C4)
    files['xl/styles.xml'] = files['xl/styles.xml']!
      .replace(/<fonts count="2">(.*?)<\/fonts>/, '<fonts count="3">$1<font><b/><sz val="11"/><color theme="4"/><name val="Calibri"/></font></fonts>')
      .replace(/<cellXfs count="2">(.*?)<\/cellXfs>/, '<cellXfs count="4">$1<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>')
    files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml']!.replace('<c r="A1"', '<c r="A1" s="2"').replace('<c r="A2"', '<c r="A2" s="3"')
    files['xl/theme/theme1.xml'] = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T"><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme></a:themeElements></a:theme>'
    files['xl/_rels/workbook.xml.rels'] = files['xl/_rels/workbook.xml.rels']!.replace('</Relationships>', `<Relationship Id="rIdT" Type="${RELS}/theme" Target="theme/theme1.xml"/></Relationships>`)
    files['[Content_Types].xml'] = files['[Content_Types].xml']!.replace('</Types>', '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
    const file = storedZip(files)
    const { book, cached } = await openXlsx(file, { keep: true })
    const s = source(book, { session: { cached } })
    expect(await s.structure({ op: 'insertRows', sheet: 'S', row: 1, count: 2 })).toMatchObject({ ok: true })
    const out = unzipText((await saveXlsx(s, { over: file })).data)
    expect(xfOf(out, 'A3')).toContain('quotePrefix="1"')
    expect(xfOf(out, 'A4')).toContain('fontId="2"')
    expect(out['xl/styles.xml']).toContain('<color theme="4"/>')
  })

  it("rejects a change the file cannot take with xlcalc's KeptRefusal and its code (0.61.0): a link removed", async () => {
    const linked = source({ sheets: [{ name: 'S', cells: { A1: { value: 'x' } } }] })
    expect(await linked.setExternalBooks([{ target: 'Report.xlsx', sheets: [{ name: 'Sheet', cells: { I8: 42 } }] }])).toMatchObject({ ok: true })
    await linked.edit([{ sheet: 'S', cell: 'B1', text: "='[Report.xlsx]Sheet'!$I$8*2" }])
    const file = (await saveXlsx(linked)).data
    const { book, cached } = await openXlsx(file, { keep: true })
    const s = source(book, { session: { cached } })
    expect(await s.setExternalBooks([])).toMatchObject({ ok: true })
    const refusal = await saveXlsx(s, { over: file }).catch((e: unknown) => e)
    expect(refusal).toMatchObject({ name: 'KeptRefusal', code: 'links' })
  })

  it('writes the links a host gave over the file (0.62.0): a new link with its values, which a file read back has', async () => {
    const file = (await saveXlsx(source({ sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }))).data
    const { book, cached } = await openXlsx(file, { keep: true })
    const s = source(book, { session: { cached } })
    await s.setExternalBooks([{ target: 'Rates.xlsx', sheets: [{ name: 'R', cells: { A1: 3 } }] }])
    await s.edit([{ sheet: 'S', cell: 'B1', text: "='[Rates.xlsx]R'!A1+1" }])
    const saved = (await saveXlsx(s, { over: file })).data
    expect(Object.keys(unzipText(saved))).toContain('xl/externalLinks/externalLink1.xml')
    const back = await openXlsx(saved)
    // entering a formula over a link marks its sheet's values as not updated, as Excel does (`refreshError`)
    expect(back.book.externalBooks).toEqual([{ target: 'Rates.xlsx', sheets: [{ name: 'R', cells: { A1: 3 }, refreshError: true }] }])
    expect(back.book.sheets[0]!.cells.B1).toMatchObject({ formula: '[1]R!A1+1' })
  })

  it("stores the calculated values over a file written without them only with writeValues (0.81.0)", async () => {
    const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' } } }] }
    const file = (await writeXlsx({ book }, { values: 'omit' })).data
    const opened = await openXlsx(file, { keep: true })
    const s = source(opened.book, { session: { cached: opened.cached } })
    expect(unzipText((await saveXlsx(s, { over: file })).data)['xl/worksheets/sheet1.xml']).not.toContain('<v>20</v>')
    const saved = (await saveXlsx(s, { over: file, writeValues: true })).data
    expect(unzipText(saved)['xl/worksheets/sheet1.xml']).toMatch(/<f>A1\*10<\/f><v>20<\/v>/)
  })
})
