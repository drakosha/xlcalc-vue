// memorySource beside cell values: Clear Formats and Clear All, hyperlinks, notes, defined names, spelling, and the default font and zoom of its sheets.
import { describe, expect, it } from 'vitest'
import { memorySource } from '../src'
import type { SheetWindow } from '../src/contract/source'
import type { SpellChecker } from '../src/contract/workbook'

const book = () =>
  memorySource({
    styles: [{}, { font: { bold: true } }],
    names: [{ name: 'Rate', ref: 'Sheet1!$A$1' }],
    sheets: [
      {
        name: 'Sheet1',
        font: { name: 'Calibri', size: 11 },
        zoom: 150,
        cells: { A1: { value: 'bold', style: 1, numberFormat: '0%' }, A2: { value: 'noted', note: { text: 'Check' } }, A3: 'plain' },
      },
    ],
  })

async function win(s: ReturnType<typeof book>): Promise<SheetWindow> {
  const w = await s.window({ sheet: 'Sheet1', row: 1, rows: 5, col: 1, cols: 3 })
  if ('pending' in w) throw new Error('pending')
  return w
}

describe('memorySource beside cell values', () => {
  it("gives a sheet's default font and zoom", async () => {
    expect((await book().sheets())[0]).toMatchObject({ font: { name: 'Calibri', size: 11 }, zoom: 150 })
  })

  it('Clear Formats keeps the value; Clear All removes the cell with its note', async () => {
    const s = book()
    await s.clear!('Sheet1', ['A1'], 'formats')
    let w = await win(s)
    expect([w.text[0]![0], w.styleIds['A1'], w.numberFormats?.['A1']]).toEqual(['bold', undefined, undefined])
    expect(await s.clear!('Sheet1', ['A2:A3'], 'all')).toMatchObject({ ok: true, changes: [{ cell: 'A2', text: '' }, { cell: 'A3', text: '' }] })
    w = await win(s)
    expect([w.text[1]![0], w.notes?.['A2']]).toEqual(['', undefined])
  })

  it('Insert Link over a range and Remove Hyperlinks; New Note and Delete Note', async () => {
    const s = book()
    await s.link!('Sheet1', 'B1:B2', { location: 'Sheet1!A3', tooltip: 'Go' })
    expect((await win(s)).links).toEqual({ B1: { location: 'Sheet1!A3', tooltip: 'Go' }, B2: { location: 'Sheet1!A3', tooltip: 'Go' } })
    await s.link!('Sheet1', 'B1:B2', null)
    expect((await win(s)).links).toEqual({})
    expect(await s.link!('Sheet1', 'B1', {})).toMatchObject({ ok: false, reason: 'refused' })
    await s.note!('Sheet1', 'c1', { author: 'Ann', text: 'New' })
    await s.note!('Sheet1', 'A2', null)
    expect((await win(s)).notes).toEqual({ C1: { author: 'Ann', text: 'New' } })
  })

  it('lists and replaces defined names, refusing names Excel rejects and duplicates', async () => {
    const s = book()
    expect(await s.names!()).toEqual([{ name: 'Rate', ref: 'Sheet1!$A$1' }])
    for (const name of ['1abc', 'B2', 'R', 'a b']) expect(await s.setNames!([{ name, ref: '1' }])).toMatchObject({ ok: false, reason: 'refused' })
    expect(await s.setNames!([{ name: 'X', ref: '1' }, { name: 'x', ref: '2' }])).toMatchObject({ ok: false })
    expect(await s.setNames!([{ name: 'X', ref: '1' }, { name: 'x', ref: '2', scope: 'Sheet1' }])).toMatchObject({ ok: true })
    expect((await s.names!()).map((n) => n.name)).toEqual(['X', 'x'])
  })

  it('spells from the active cell on and wraps around, with Add to Dictionary going to the checker', async () => {
    const added: string[] = []
    const known = new Set(['bold', 'noted', 'plain', 'good'])
    const checker: SpellChecker = { check: (w) => known.has(w.toLowerCase()), suggest: () => ['good'], add: (w) => added.push(w) }
    const s = memorySource({ sheets: [{ name: 'S', cells: { A1: 'good bda', A2: { formula: '"zzz"' }, A3: 'NASA wrod' } }] }, { spelling: checker })
    expect(await s.spelling!({ sheet: 'S', after: 'A2' })).toEqual({ cell: 'A3', word: 'wrod', at: 5, suggestions: ['good'] })
    expect(await s.spelling!({ sheet: 'S', after: 'A3', from: 6 })).toMatchObject({ cell: 'A1', word: 'bda', at: 5 })
    expect(await s.spelling!({ sheet: 'S', ignore: ['BDA', 'wrod'] })).toBeNull()
    await s.learnWord!('bda')
    expect(added).toEqual(['bda'])
    expect(memorySource({ sheets: [] }).spelling).toBeUndefined()
  })
})
