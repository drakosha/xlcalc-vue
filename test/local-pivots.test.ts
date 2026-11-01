// @vitest-environment node
// localSource over a workbook with a PivotTable: its areas among the sheet's information, Excel's alerts for what its cells
// do not take, its labels renamed by typing, Refresh and Refresh All, and a protected sheet's Use PivotTable and PivotChart.
import { afterEach, describe, expect, it } from 'vitest'
import type { EditOutcome, SheetOp } from '../src/contract/source'
import { closeAll, source, win } from './local-fixture'
import { filteredPivotBook, pivotBook } from './pivot-fixture'

afterEach(closeAll)

// Excel 16.0 for Windows' answers to xlcalc's set `edit`
const PART = "We can't change this part of the PivotTable."
const AFFECT = "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again."
const FORMULA = 'Cannot enter a formula for an item or field name in a PivotTable report.'
const PROTECTED = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password."
// Excel 16.0 for Windows' alerts for Refresh All (Ctrl+Alt+F5) on a protected sheet, without and with Use PivotTable and PivotChart
const PIVOT_PROTECTED = 'Cannot edit PivotTable on protected sheet.'
const PIVOT_SHARED =
  'That command cannot be performed while a protected sheet contains another PivotTable report based on the same source data.\n\nTo remove protection from the sheet that has the other report, click the sheet tab, and then click Unprotect Sheet (Review tab, Protect group). Then try the command again.'
const refusedWith = (message: string) => ({ ok: false, reason: 'refused', code: 'pivot', message })
const ok = (out: EditOutcome) => expect(out).toMatchObject({ ok: true })
const valueAt = async (s: ReturnType<typeof source>, cell: string) => {
  const m = /^([A-Z])(\d+)$/.exec(cell)!
  return (await win(s, 'Report', Number(m[2]), 1, m[1]!.charCodeAt(0) - 64, 1)).values[0]![0]
}

describe('localSource: PivotTables', () => {
  it("lists each sheet's PivotTables with their areas", async () => {
    expect((await source(pivotBook()).sheets()).map((x) => x.pivots)).toEqual([undefined, [{ name: 'PivotTable1', range: 'A3:B6', headers: 'A3:B3', rowLabels: 'A4:A6', values: 'B4:B6', cache: 1, fields: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Kind' }, { name: 'Year' }] }]])
    expect((await source(filteredPivotBook()).sheets())[1]!.pivots).toMatchObject([{ name: 'PivotTable1', range: 'A3:D7', filters: 'A1:B1', headers: 'A3:D4', rowLabels: 'A5:A7', values: 'B5:D7' }])
  })

  it("reads a PivotTable's definition with its cache's fields, without the saved lines and rows; null for none", async () => {
    const s = source(pivotBook())
    const got = await s.pivotTable('report', 'pivottable1')
    expect(got?.table).toMatchObject({ name: 'PivotTable1', rows: [0], values: [{ field: 1, name: 'Sum of Amount' }] })
    expect(got?.table).not.toHaveProperty('rowLines')
    expect(got?.cache.fields.map((f) => f.name)).toEqual(['Region', 'Amount', 'Kind', 'Year'])
    expect(got?.cache).not.toHaveProperty('records')
    expect(await s.pivotTable('Report', 'PivotTable9')).toBeNull()
    expect(await s.pivotTable('Data', 'PivotTable1')).toBeNull()
  })

  it('lists a PivotTable not laid out, or shrunk to its top-left cell (#SPILL!), by its table and values alone', async () => {
    const fields = [{ name: 'Region' }, { name: 'Amount' }, { name: 'Kind' }, { name: 'Year' }]
    // added without the engine: Excel's empty PivotTable of three columns and 18 rows, without saved lines
    const placeholder = pivotBook({ table: { location: { range: 'A3:C20', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 0 }, rowLines: undefined, columnLines: undefined } })
    expect((await source(placeholder).sheets())[1]!.pivots).toEqual([{ name: 'PivotTable1', range: 'A3:C20', values: 'A4:C20', cache: 1, fields }])
    const spilled = pivotBook({ table: { location: { range: 'A3', firstHeaderRow: 0, firstDataRow: 0, firstDataColumn: 0 } } })
    expect((await source(spilled).sheets())[1]!.pivots).toEqual([{ name: 'PivotTable1', range: 'A3', values: 'A3', cache: 1, fields }])
  })

  it("refuses a value or a formula typed, pasted or filled over its cells with Excel's alert; a formula has its own", async () => {
    const s = source(pivotBook())
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: '100' }])).toEqual(refusedWith(PART))
    // a formula sent over a value: Excel's object model refuses it as a formula (xlcalc's set alerts-held); the sheet's own typing never sends it
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: '=1+1' }])).toEqual(refusedWith(FORMULA))
    expect(await s.edit([{ sheet: 'Report', cell: 'A4', text: '="East"' }])).toEqual(refusedWith(FORMULA))
    // texts entered into several cells at once, and a paste, affect it (xlcalc's sets alerts and alerts-held)
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: '1' }, { sheet: 'Report', cell: 'C4', text: '2' }])).toEqual(refusedWith(AFFECT))
    // a formula over an empty header cell, which holds no label, is a formula over its cells all the same
    expect(await source(filteredPivotBook()).edit([{ sheet: 'Report', cell: 'C3', text: '=1' }])).toEqual(refusedWith(FORMULA))
    expect(await s.paste('Report', 'B5', '7\t8')).toEqual(refusedWith(AFFECT))
    // Ctrl+Enter with the active cell in it enters that cell alone, a value over a value
    expect(await s.fill('Report', 'B4:C5', '9', 'B4')).toEqual(refusedWith(PART))
    expect(await valueAt(s, 'B4')).toBe(15)
    // the cells beside it take anything
    ok(await s.edit([{ sheet: 'Report', cell: 'C4', text: '2' }]))
  })

  it("refuses clearing its cells and an operation that would cut it with Excel's alert for a change that affects it", async () => {
    const s = source(pivotBook())
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: null }])).toEqual(refusedWith(AFFECT))
    expect(await s.clear('Report', ['A4:B5'], 'all')).toEqual(refusedWith(AFFECT))
    expect(await s.structure({ op: 'insertRows', sheet: 'Report', row: 5, count: 1 })).toEqual(refusedWith(AFFECT))
    expect(await s.structure({ op: 'deleteColumns', sheet: 'Report', column: 'B', count: 1 })).toEqual(refusedWith(AFFECT))
    expect(await valueAt(s, 'B6')).toBe(35)
    // an operation on the PivotTable keeps the library's reason
    expect(await s.structure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'Nope' }] })).toMatchObject({ ok: false, code: 'pivot', message: expect.stringMatching(/Nope/) })
  })

  it('deletes a PivotTable whose cells are all cleared, as Excel does, in one undo step, and reads the sheets again', async () => {
    const s = source(pivotBook())
    const events: string[] = []
    s.subscribe((e) => events.push(e.kind))
    const cells = ['A3', 'B3', 'A4', 'B4', 'A5', 'B5', 'A6', 'B6'].map((cell) => ({ sheet: 'Report', cell, text: null }))
    ok(await s.edit(cells))
    expect(events).toContain('reload')
    expect((await s.sheets())[1]!.pivots).toBeUndefined()
    expect(await valueAt(s, 'B6')).toBeNull()
    ok(await s.undo())
    expect((await s.sheets())[1]!.pivots).toMatchObject([{ name: 'PivotTable1' }])
    ok(await s.clear('Report', ['A3:B6'], 'all'))
    expect((await s.sheets())[1]!.pivots).toBeUndefined()
  })

  it("keeps xlcalc's own reason where Excel's alert is not one of those: a name taken, an item a filter does not have", async () => {
    const s = source(pivotBook())
    expect(await s.edit([{ sheet: 'Report', cell: 'B3', text: 'Region' }])).toEqual(refusedWith('PivotTable field name already exists: "Region"'))
    expect(await source(filteredPivotBook()).edit([{ sheet: 'Report', cell: 'B1', text: '2030' }])).toEqual(refusedWith('"2030" is not an item of this field'))
  })

  it('refuses a range copied or moved onto it, and a move of part of it, with the alert that it affects a PivotTable', async () => {
    const s = source(pivotBook({ extra: { E1: { value: 1 }, F1: { value: 2 } } }))
    expect(await s.structure({ op: 'copyRange', from: { sheet: 'Report', range: 'E1:F1' }, to: { sheet: 'Report', range: 'A5' } })).toEqual(refusedWith(AFFECT))
    expect(await s.structure({ op: 'moveRange', from: { sheet: 'Report', range: 'E1:F1' }, to: { sheet: 'Report', cell: 'A5' } })).toEqual(refusedWith(AFFECT))
    expect(await s.structure({ op: 'moveRange', from: { sheet: 'Report', range: 'A4:B4' }, to: { sheet: 'Report', cell: 'E4' } })).toEqual(refusedWith(AFFECT))
  })

  it("words the alerts in the workbook's language", async () => {
    const s = source(pivotBook({ locale: 'ru-RU' }))
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: '100' }])).toMatchObject({ code: 'pivot', message: 'Мы не можем изменить эту часть сводной таблицы.' })
    expect(await s.edit([{ sheet: 'Report', cell: 'A4', text: '=1+1' }])).toMatchObject({ code: 'pivot', message: 'Нельзя задавать формулу для имени поля или элемента в отчете сводной таблицы.' })
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: null }])).toMatchObject({
      code: 'pivot',
      message: 'Мы не можем изменить выделенные ячейки, так как это повлияет на сводную таблицу. Для изменения отчета используйте список полей. Если вы пытаетесь вставить или удалить ячейки, переместить сводную таблицу и повторите попытку.',
    })
  })

  it('renames a label typed over, as Excel does, in one undo step', async () => {
    const s = source(pivotBook())
    ok(await s.edit([{ sheet: 'Report', cell: 'A4', text: 'Nord' }]))
    expect(await valueAt(s, 'A4')).toBe('Nord')
    ok(await s.edit([{ sheet: 'Report', cell: 'B3', text: 'Total amount' }]))
    expect(await valueAt(s, 'B3')).toBe('Total amount')
    ok(await s.undo())
    ok(await s.undo())
    expect([await valueAt(s, 'A4'), await valueAt(s, 'B3')]).toEqual(['North', 'Sum of Amount'])
  })

  it('refreshes the PivotTable from its source as one undo step, and Refresh All refreshes every PivotTable', async () => {
    const s = source(pivotBook())
    ok(await s.edit([{ sheet: 'Data', cell: 'B3', text: '200' }]))
    ok(await s.structure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] }))
    expect(await valueAt(s, 'B5')).toBe(200)
    ok(await s.undo())
    expect(await valueAt(s, 'B5')).toBe(20)
    // a workbook with PivotTables has Refresh All without calls outside it to ask
    expect(s.refreshConnections).toBeTypeOf('function')
    ok(await s.refreshConnections!())
    expect(await valueAt(s, 'B6')).toBe(215)
    // Refresh All of some functions' calls leaves the PivotTables
    ok(await s.edit([{ sheet: 'Data', cell: 'B4', text: '0' }]))
    ok(await s.refreshConnections!(['WEBSERVICE']))
    expect(await valueAt(s, 'B6')).toBe(215)
  })

  it('has Refresh All in every workbook: without PivotTables or calls outside it, it changes nothing', async () => {
    expect(await source({ sheets: [{ name: 'S', cells: {} }] }).refreshConnections()).toEqual({ ok: true, changes: [] })
  })

  it('changes a PivotTable on a protected sheet only with Use PivotTable and PivotChart, as Microsoft documents the box', async () => {
    const locked = source(pivotBook({ protection: {} }))
    const refused = { ok: false, reason: 'refused', code: 'protected', message: PROTECTED }
    expect(await locked.structure({ op: 'editPivotTable', sheet: 'Report', pivot: 'PivotTable1', set: { grandTotals: { rows: false } } })).toEqual(refused)
    expect(await locked.structure({ op: 'addPivotTable', sheet: 'Report', cell: 'F3', source: { type: 'range', sheet: 'Data', range: 'A1:D4' } })).toEqual(refused)
    const allowed = source(pivotBook({ protection: { pivotTables: true } }))
    ok(await allowed.structure({ op: 'editPivotTable', sheet: 'Report', pivot: 'PivotTable1', set: { grandTotals: { rows: false } } }))
    expect(await valueAt(allowed, 'A6')).toBeNull()
  })

  it("refuses Refresh and Refresh All on a protected sheet whatever it allows, with Excel for Windows' alerts", async () => {
    const refresh: SheetOp = { op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] }
    const refusedBy = (message: string) => ({ ok: false, reason: 'refused', code: 'protected', message })
    const locked = source(pivotBook({ protection: {} }))
    ok(await locked.edit([{ sheet: 'Data', cell: 'B2', text: '15' }]))
    expect(await locked.structure(refresh)).toEqual(refusedBy(PIVOT_PROTECTED))
    expect(await locked.refreshConnections!()).toEqual(refusedBy(PIVOT_PROTECTED))
    expect(await valueAt(locked, 'B4')).toBe(15)
    const allowed = source(pivotBook({ protection: { pivotTables: true } }))
    ok(await allowed.edit([{ sheet: 'Data', cell: 'B2', text: '15' }]))
    expect(await allowed.structure(refresh)).toEqual(refusedBy(PIVOT_SHARED))
    expect(await allowed.refreshConnections!()).toEqual(refusedBy(PIVOT_SHARED))
    expect(await valueAt(allowed, 'B4')).toBe(15)
    const ru = source(pivotBook({ locale: 'ru-RU', protection: {} }))
    expect(await ru.structure(refresh)).toEqual(refusedBy('Нельзя изменять "сводной таблицы" на защищенном листе.'))
    expect(await source(pivotBook({ locale: 'ru-RU', protection: { pivotTables: true } })).refreshConnections!()).toMatchObject({ message: expect.stringMatching(/^Эта команда не может быть выполнена, пока защищенный лист/) })
  })

  it('renames an unlocked label on a protected sheet only with Use PivotTable reports', async () => {
    const unlocked = { styles: [{ protection: { locked: false } }], extra: { A4: { value: 'North', style: 0 } } }
    const locked = source(pivotBook({ ...unlocked, protection: {} }))
    expect(await locked.edit([{ sheet: 'Report', cell: 'A4', text: 'Nord' }])).toEqual({ ok: false, reason: 'refused', code: 'protected', message: PROTECTED })
    const allowed = source(pivotBook({ ...unlocked, protection: { pivotTables: true } }))
    ok(await allowed.edit([{ sheet: 'Report', cell: 'A4', text: 'Nord' }]))
    expect(await valueAt(allowed, 'A4')).toBe('Nord')
  })

  it('gates slicers on a protected sheet: their choice by Use PivotTable reports, adding and deleting by Edit objects too', async () => {
    const refused = { ok: false, reason: 'refused', code: 'protected', message: PROTECTED }
    const slicer: SheetOp = { op: 'addSlicer', sheet: 'Report', slicer: {}, cache: { field: 'Region', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] } }
    const choose: SheetOp = { op: 'editSlicer', slicer: 'Region', cache: { items: [{ item: 0, selected: true }, { item: 1, selected: false }] } }
    const protect = async (box: object) => {
      const s = source(pivotBook())
      ok(await s.structure(slicer))
      ok(await s.protect('Report', box))
      return s
    }
    const none = await protect({})
    expect(await none.structure(choose)).toEqual(refused)
    expect(await none.structure({ op: 'deleteSlicer', slicer: 'Region' })).toEqual(refused)
    const pivots = await protect({ pivotTables: true })
    ok(await pivots.structure(choose))
    expect(await valueAt(pivots, 'A5')).toBe('Grand Total')
    expect(await pivots.structure({ op: 'deleteSlicer', slicer: 'Region' })).toEqual(refused)
    expect(await pivots.structure({ ...slicer, slicer: { name: 'Second' } } as SheetOp)).toEqual(refused)
    const both = await protect({ pivotTables: true, objects: true })
    ok(await both.structure({ op: 'deleteSlicer', slicer: 'Region' }))
  })
})

describe('localSource: PivotTables and AutoFill', () => {
  it('keeps every PivotTable of the workbook when a range is filled elsewhere', async () => {
    const s = source(pivotBook({ extra: { E1: { value: 1 }, E2: { value: 2 } } }))
    expect(await s.structure({ op: 'fillRange', sheet: 'Report', range: 'E1:E2', to: 'E1:E5' })).toMatchObject({ ok: true })
    expect((await s.sheets()).find((sh) => sh.name === 'Report')?.pivots?.map((p) => p.name)).toEqual(['PivotTable1'])
  })
})
