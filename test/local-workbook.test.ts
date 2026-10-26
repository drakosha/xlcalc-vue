// @vitest-environment node
// localSource beside cell values: Clear Formats and Clear All, hyperlinks, notes, defined names, rules, tables, dependents and spelling, with formulas in the workbook's locale.
import { afterEach, describe, expect, it } from 'vitest'
import type { SourceEvent } from '../src/contract/source'
import type { SpellChecker } from '../src/contract/workbook'
import { BOLD, book, closeAll, source, win } from './local-fixture'

afterEach(closeAll)

/** A checker that knows these words, suggests `fix` for any other, and keeps the words added. */
function checker(...known: string[]): SpellChecker & { added: string[] } {
  const words = new Set(known.map((w) => w.toLowerCase()))
  const added: string[] = []
  return { added, check: (w) => words.has(w.toLowerCase()), suggest: () => ['fix'], add: (w) => (words.add(w.toLowerCase()), added.push(w)) }
}

describe('localSource: clearing', () => {
  it('Clear Formats removes the style and keeps the value; Clear All removes the value, the note and the hyperlink, as one undo step each', async () => {
    const s = source()
    expect(await s.clear('Лист1', ['C1'], 'formats')).toMatchObject({ ok: true })
    let w = await win(s, 'Лист1', 1, 5, 1, 5)
    expect(w.text[0]![2]).toBe('жирный')
    expect(w.styles[w.styleIds['C1'] ?? 0]).not.toMatchObject(BOLD)
    expect(await s.clear('Лист1', ['B1:B2'], 'all')).toMatchObject({ ok: true })
    w = await win(s, 'Лист1', 1, 5, 1, 5)
    expect([w.text[0]![1], w.text[1]![1], w.notes?.['B1'], w.links?.['B2']]).toEqual(['', '', undefined, undefined])
    await s.undo()
    w = await win(s, 'Лист1', 1, 5, 1, 5)
    expect(w.notes?.['B1']?.text).toBe('Проверить')
  })
})

describe('localSource: hyperlinks and notes', () => {
  it('Insert Link over a range sets each cell, Remove Hyperlinks removes them, and the windows over them are repainted', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    expect(await s.link('Лист1', 'A1:A2', { href: 'https://example.com/', tooltip: 'Пример' })).toMatchObject({ ok: true })
    let w = await win(s, 'Лист1', 1, 3, 1, 3)
    expect(w.links?.['A1']).toMatchObject({ href: 'https://example.com/', tooltip: 'Пример' })
    expect(w.links?.['A2']).toMatchObject({ href: 'https://example.com/' })
    expect(events).toContainEqual({ kind: 'repaint', ranges: [{ sheet: 'Лист1', range: 'A1:A2' }] })
    await s.link('Лист1', 'A1:A2', null)
    w = await win(s, 'Лист1', 1, 3, 1, 3)
    expect(w.links?.['A1']).toBeUndefined()
  })

  it('refuses a hyperlink with neither an address nor a place', async () => {
    expect(await source().link('Лист1', 'A1', {})).toMatchObject({ ok: false })
  })

  it('New Note, Edit Note and Delete Note on a cell', async () => {
    const s = source()
    await s.note('Лист1', 'a2', { author: 'Михаил', text: 'Новое' })
    expect((await win(s, 'Лист1', 1, 3, 1, 3)).notes?.['A2']).toEqual({ author: 'Михаил', text: 'Новое' })
    await s.note('Лист1', 'B1', null)
    expect((await win(s, 'Лист1', 1, 3, 1, 3)).notes?.['B1']).toBeUndefined()
  })
})

describe('localSource: defined names', () => {
  it('lists names with their formulas in the workbook locale, and a name set in ru-RU calculates in a formula', async () => {
    const s = source({ ...book(), names: [{ name: 'Итого', ref: "SUM('Лист1'!$A$1:$A$2,0.5)" }] })
    expect(await s.names()).toEqual([{ name: 'Итого', ref: 'СУММ(Лист1!$A$1:$A$2;0,5)' }])
    expect(await s.setNames([{ name: 'Двойка', ref: 'Лист1!$A$2' }, { name: 'Мест', ref: '1,5', scope: 'Лист1' }])).toMatchObject({ ok: true })
    expect(await s.names()).toEqual([
      { name: 'Двойка', ref: 'Лист1!$A$2' },
      { name: 'Мест', ref: '1,5', scope: 'Лист1' },
    ])
    const got = await s.edit([{ sheet: 'Лист1', cell: 'H1', text: '=Двойка+Мест' }])
    expect(got.ok && got.changes.find((c) => c.cell === 'H1')?.value).toBe(3.5)
  })

  it("does not list a hidden name, as Excel's Name Manager does not, and keeps it when the list is replaced", async () => {
    const s = source({ ...book(), names: [{ name: 'Seen', ref: "'Лист1'!$A$1" }, { name: '_Hidden', ref: "'Лист1'!$A$2", hidden: true }] })
    expect((await s.names()).map((n) => n.name)).toEqual(['Seen'])
    expect(await s.setNames([{ name: 'Other', ref: 'Лист1!$A$3' }])).toMatchObject({ ok: true })
    expect((await s.names()).map((n) => n.name)).toEqual(['Other'])
    expect((await s.book()).names).toEqual([
      { name: 'Other', ref: "'Лист1'!$A$3" },
      { name: '_Hidden', ref: "'Лист1'!$A$2", hidden: true },
    ])
  })

  it('refuses a formula that does not parse in the workbook locale and a name the library rejects, keeping the names', async () => {
    const s = source()
    expect(await s.setNames([{ name: 'Плохое', ref: 'СУММ(' }])).toMatchObject({ ok: false, reason: 'refused' })
    for (const name of ['1abc', 'Q1', 'ИСТИНА']) expect(await s.setNames([{ name, ref: '1' }])).toMatchObject({ ok: false, reason: 'refused' })
    expect(await s.setNames([{ name: 'X', ref: '1' }, { name: 'x', ref: '2' }])).toMatchObject({ ok: false, reason: 'refused' })
    expect(await s.names()).toEqual([])
  })
})

describe('localSource: dependents', () => {
  it('the cells whose formulas refer to a cell', async () => {
    expect(await source().dependents('Лист1', 'A1')).toEqual({ cells: [{ sheet: 'Лист1', cell: 'A3' }], more: false })
  })
})

describe('localSource: rules', () => {
  it('reads and writes conditional formats and data validation with their formulas in the workbook locale', async () => {
    const s = source()
    expect(await s.rules('Лист1')).toEqual({ conditionalFormats: [], validations: [] })
    const set = await s.setRules('Лист1', {
      conditionalFormats: [{ type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['1,5'], style: { font: { bold: true } } }],
      validations: [{ range: 'D1:D9', type: 'decimal', operator: 'between', formula1: '0', formula2: 'СУММ($A$1;0,5)' }],
    })
    expect(set).toMatchObject({ ok: true })
    const got = await s.rules('Лист1')
    expect(got.conditionalFormats[0]).toMatchObject({ type: 'cellIs', formulas: ['1,5'] })
    expect(got.validations[0]).toMatchObject({ formula1: '0', formula2: 'СУММ($A$1;0,5)' })
    const stored = (await s.book()).sheets[0]!
    expect(stored.conditionalFormats?.[0]).toMatchObject({ formulas: ['1.5'] })
    expect(stored.validations?.[0]).toMatchObject({ formula2: 'SUM($A$1,0.5)' })
  })
})

describe('localSource: tables', () => {
  it('Format as Table creates a table, Convert to Range removes it', async () => {
    const s = source()
    const add = await s.table({ op: 'addTable', sheet: 'Список', table: { name: 'Люди', range: 'A1:B6', columns: [{ name: 'Имя' }, { name: 'Число' }] } })
    expect(add).toMatchObject({ ok: true })
    expect((await win(s, 'Список', 1, 6, 1, 2)).tables).toEqual([{ name: 'Люди', range: 'A1:B6', headerRow: 1 }])
    expect(await s.table({ op: 'convertTable', table: 'Люди' })).toMatchObject({ ok: true })
    expect((await win(s, 'Список', 1, 6, 1, 2)).tables).toEqual([])
  })
})

describe('localSource: spelling', () => {
  it('has no spelling without a checker', () => {
    expect(source().spelling).toBeUndefined()
  })

  it("finds misspelt words from the active cell on, skipping formulas, uppercase words, words with digits and Ignore All's words", async () => {
    const c = checker('good', 'text')
    const s = source({ sheets: [{ name: 'S', cells: { A1: { value: 'good tekst' }, A2: { value: 'NASA x1 good' }, B2: { formula: '"wrng"' }, A3: { value: 'bad wurd' } } }] }, { spelling: c })
    expect(await s.spelling!({ sheet: 'S' })).toEqual({ cell: 'A1', word: 'tekst', at: 5, suggestions: ['fix'] })
    expect(await s.spelling!({ sheet: 'S', after: 'A1', from: 6 })).toMatchObject({ cell: 'A3', word: 'bad', at: 0 })
    expect(await s.spelling!({ sheet: 'S', after: 'A3', from: 1, ignore: ['Wurd', 'tekst'] })).toBeNull()
    await s.learnWord!('tekst')
    expect(c.added).toEqual(['tekst'])
  })
})
