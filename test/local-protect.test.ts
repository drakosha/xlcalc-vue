// @vitest-environment node
// localSource on a protected sheet: what it refuses and allows, as Excel for Mac 16.114 refused through its object model.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { EditOutcome } from '../src/contract/source'
import { closeAll, source, styleOf, win } from './local-fixture'

afterEach(closeAll)

const ALERT_EN = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password."
const ALERT_RU = 'Ячейка или диаграмма, которую вы пытаетесь изменить, находится на защищенном листе. Чтобы внести изменения, снимите защиту листа. Может потребоваться пароль.'

/** A form: A column locked, B1:B3 unlocked, row 10 unlocked entire, A2's formula hidden; `Free` unprotected. */
const form = (protection: BookInput['sheets'][number]['protection'] = {}, locale = 'en-US'): BookInput => ({
  locale,
  styles: [{}, { protection: { locked: false } }, { protection: { hidden: true } }],
  sheets: [
    {
      name: 'Form',
      cells: { A1: { value: 'Name' }, B1: { value: 'Ann', style: 1 }, A2: { formula: 'LEN(B1)', style: 2 }, B2: { value: 'hello', style: 1 }, B3: { value: 'b', style: 1 } },
      rows: { '10': { style: 1 } },
      protection,
    },
    { name: 'Free', cells: { A1: { value: 1 } } },
  ],
})
const refused = { ok: false, reason: 'refused', code: 'protected', message: ALERT_EN }
const ok = (out: EditOutcome) => expect(out).toMatchObject({ ok: true })

describe('localSource: sheet protection', () => {
  it('lists a protected sheet with what it allows; protect and unprotect are one undo step each', async () => {
    const s = source(form({ formatCells: true }))
    expect((await s.sheets()).map((x) => x.protection)).toEqual([{ formatCells: true }, undefined])
    ok(await s.edit([{ sheet: 'Free', cell: 'A1', text: '2' }]))
    ok(await s.protect('Free', {}))
    expect((await s.sheets())[1]!.protection).toEqual({})
    expect(await s.edit([{ sheet: 'Free', cell: 'A1', text: '3' }])).toEqual(refused)
    ok(await s.protect('Form', null))
    expect((await s.sheets())[0]!.protection).toBeUndefined()
    ok(await s.edit([{ sheet: 'Form', cell: 'A1', text: 'x' }]))
    await s.undo()
    await s.undo()
    await s.undo()
    expect((await s.sheets()).map((x) => x.protection)).toEqual([{ formatCells: true }, undefined])
  })

  it("refuses Data > Filter on a protected sheet whatever its boxes allow; an unprotected sheet takes it", async () => {
    const s = source(form({ autoFilter: true, sort: true }))
    expect(await s.setAutoFilter('Form', 'A1:B3')).toEqual(refused)
    expect((await s.sheets())[0]!.autoFilter).toBeUndefined()
    ok(await s.setAutoFilter('Free', 'A1'))
  })

  it("windows carry each cell's protection in its style", async () => {
    const w = await win(source(form()), 'Form', 1, 2, 1, 2)
    expect(styleOf(w, 'B1')?.protection).toEqual({ locked: false })
    expect(styleOf(w, 'A2')?.protection).toEqual({ hidden: true })
  })

  it("refuses typing into a locked cell with Excel's alert, in the workbook's language; unlocked cells take it", async () => {
    const s = source(form())
    expect(await s.edit([{ sheet: 'Form', cell: 'A1', text: 'x' }])).toEqual(refused)
    ok(await s.edit([{ sheet: 'Form', cell: 'B1', text: 'Bob' }]))
    // a locked cell anywhere among the entries refuses them all, and nothing is written
    expect(await s.edit([{ sheet: 'Form', cell: 'B2', text: 'y' }, { sheet: 'Form', cell: 'A1', text: 'x' }])).toEqual(refused)
    expect((await s.book()).sheets[0]!.cells.B2?.value).toBe('hello')
    // an empty cell is locked unless its row, column or own style unlocks it
    expect(await s.edit([{ sheet: 'Form', cell: 'Z100', text: '1' }])).toEqual(refused)
    ok(await s.edit([{ sheet: 'Form', cell: 'X10', text: '1' }]))
    ok(await s.edit([{ sheet: 'Free', cell: 'A1', text: '2' }]))
    const ru = source(form({}, 'ru-RU'))
    expect(await ru.edit([{ sheet: 'Form', cell: 'A1', text: 'x' }])).toEqual({ ...refused, message: ALERT_RU })
  })

  it('refuses a paste, a fill, a pick and clearing that reach a locked cell', async () => {
    const s = source(form())
    expect(await s.paste('Form', 'B1', 'x\ty')).toEqual(refused)
    ok(await s.paste('Form', 'B1', 'x\ny'))
    expect(await s.fill('Form', 'A1:B1', 'z', 'B1')).toEqual(refused)
    ok(await s.fill('Form', 'B1:B3', 'z', 'B1'))
    expect(await s.pick('Form', 'A1', 0, 'x')).toEqual(refused)
    expect(await s.edit([{ sheet: 'Form', cell: 'A1', text: null }])).toEqual(refused)
    expect(await s.clear('Form', ['A1:B1'], 'formats')).toEqual(refused)
  })

  it('clears unlocked cells as Excel does: Clear Formats locks them again, Clear All without Format cells keeps their formats', async () => {
    const s = source(form())
    ok(await s.clear('Form', ['B1'], 'formats'))
    ok(await s.clear('Form', ['B2'], 'all'))
    const w = await win(s, 'Form', 1, 2, 2, 1)
    expect(w.text).toEqual([['Ann'], ['']])
    expect([styleOf(w, 'B1')?.protection, styleOf(w, 'B2')?.protection]).toEqual([undefined, { locked: false }])
  })

  it('formats cells only with Format cells, locked ones too, and never changes Locked or Hidden', async () => {
    expect(await source(form()).format('Form', [{ range: 'B1', style: { font: { bold: true } } }])).toEqual(refused)
    const s = source(form({ formatCells: true }))
    ok(await s.format('Form', [{ range: 'A1', style: { font: { bold: true } } }]))
    expect(await s.format('Form', [{ range: 'B1', style: { protection: { locked: null } } }])).toEqual(refused)
    expect(await s.format('Form', [{ range: 'B1', style: null }])).toEqual(refused)
  })

  it('sizes and hides columns only with Format columns, rows only with Format rows', async () => {
    expect(await source(form()).layout({ sheet: 'Form', columns: 'A', width: 20 })).toEqual(refused)
    expect(await source(form({ formatColumns: true })).layout({ sheet: 'Form', rows: '1', height: 30 })).toEqual(refused)
    ok(await source(form({ formatColumns: true })).layout({ sheet: 'Form', columns: 'C', hidden: true }))
    ok(await source(form({ formatRows: true })).layout({ sheet: 'Form', rows: '1', height: 30 }))
  })

  it('inserts rows and columns with their boxes; deletes them only when none of their cells is locked, with the same alert', async () => {
    expect(await source(form()).structure({ op: 'insertRows', sheet: 'Form', row: 1 })).toEqual(refused)
    ok(await source(form({ insertRows: true })).structure({ op: 'insertRows', sheet: 'Form', row: 1 }))
    ok(await source(form({ insertColumns: true })).structure({ op: 'insertColumns', sheet: 'Form', column: 'A' }))
    expect(await source(form()).structure({ op: 'deleteRows', sheet: 'Form', row: 10 })).toEqual(refused)
    const rows = source(form({ deleteRows: true }))
    // Excel for Windows shows its protected sheet's alert here too, with no text of its own for a locked cell
    expect(await rows.structure({ op: 'deleteRows', sheet: 'Form', row: 1 })).toEqual(refused)
    ok(await rows.structure({ op: 'deleteRows', sheet: 'Form', row: 10 }))
    const cols = source(form({ deleteColumns: true }, 'ru-RU'))
    expect(await cols.structure({ op: 'deleteColumns', sheet: 'Form', column: 'B' })).toEqual({ ...refused, message: ALERT_RU })
    // sheets are the workbook's structure, which sheet protection does not cover
    ok(await source(form()).structure({ op: 'renameSheet', sheet: 'Form', name: 'Anketa' }))
  })

  it('pastes, moves and fills into unlocked cells only; merges nothing', async () => {
    const s = source(form())
    ok(await s.structure({ op: 'copyRange', from: { sheet: 'Free', range: 'A1' }, to: { sheet: 'Form', range: 'B1:B2' } }))
    expect(await s.structure({ op: 'copyRange', from: { sheet: 'Free', range: 'A1' }, to: { sheet: 'Form', range: 'A1' } })).toEqual(refused)
    expect(await s.structure({ op: 'copyRange', from: { sheet: 'Form', range: 'B1:B2' }, to: { sheet: 'Form', range: 'B3' } })).toEqual(refused)
    expect(await s.structure({ op: 'moveRange', from: { sheet: 'Form', range: 'A1' }, to: { sheet: 'Free', cell: 'C1' } })).toEqual(refused)
    expect(await s.structure({ op: 'fillRange', sheet: 'Form', range: 'B1', to: 'B1:B5' })).toEqual(refused)
    expect(await s.structure({ op: 'copyRange', from: { sheet: 'Free', range: 'A1' }, to: { sheet: 'Form', range: 'B1' }, paste: 'formats' })).toEqual(refused)
    expect(await s.structure({ op: 'mergeCells', sheet: 'Form', range: 'B1:B2' })).toEqual(refused)
  })

  it('inserts hyperlinks only with Insert hyperlinks and into unlocked cells, notes only with Edit objects', async () => {
    expect(await source(form()).link('Form', 'B1', { href: 'https://example.org/' })).toEqual(refused)
    const links = source(form({ insertHyperlinks: true }))
    ok(await links.link('Form', 'B1', { href: 'https://example.org/' }))
    expect(await links.link('Form', 'A1', { href: 'https://example.org/' })).toEqual(refused)
    expect(await source(form()).note('Form', 'B1', { text: 'x' })).toEqual(refused)
    ok(await source(form({ objects: true })).note('Form', 'A1', { text: 'x' }))
  })

  it('refuses rules, tables, the outline and a Replace All that reaches a locked cell', async () => {
    const s = source(form({ formatCells: true }))
    expect(await s.setRules('Form', { validations: [] })).toEqual(refused)
    ok(await s.setRules('Form', { conditionalFormats: [] }))
    expect(await source(form()).setRules('Form', { conditionalFormats: [] })).toEqual(refused)
    expect(await s.table({ op: 'addTable', sheet: 'Form', table: { name: 'T', range: 'B1:B3', columns: [{ name: 'b' }] } })).toEqual({ ...refused, message: "Table features aren't available because the sheet is protected." })
    expect(await s.outline({ op: 'group', sheet: 'Form', rows: '2:3' })).toEqual(refused)
    expect(await s.replace({ sheet: 'Form', query: 'n' }, 'm')).toEqual({ ...refused, replaced: 0 })
    expect(await s.replace({ sheet: 'Form', query: 'hello' }, 'bye')).toMatchObject({ ok: true, replaced: 1 })
  })
})
