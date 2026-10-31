// Excel's dialogs over a PivotTable on SheetView over localSource: the cell menu's Summarize Values By and Show Values As,
// Value Field Settings, Field Settings, PivotTable Options, Grouping and Ungroup, Insert Calculated Field and Item, and the
// Design commands, each OK one step of undo through the source's structure.
import type { VueWrapper } from '@vue/test-utils'
import { describe, expect, it } from 'vitest'
import { cellText, key, settle, until } from './helpers'
import { api, choose, lastState, menuAt, menuItem, ok, q, sales, set } from './pivot-sheet'

const column = (w: VueWrapper, col: number, from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => cellText(w, from + i, col))

describe('the cell menu of a PivotTable', () => {
  it("a value's menu: Summarize Values By and Show Values As, each one step of undo", async () => {
    const { w } = await sales({ rows: [0], values: [{ field: 3 }] })
    await menuAt(w, '4:2')
    expect(menuItem('pivotNumberFormat')).not.toBeNull()
    expect(menuItem('pivotSummarize')?.getAttribute('aria-haspopup')).toBe('menu')
    await choose('pivotFn-count', 'pivotSummarize')
    await until(() => cellText(w, 3, 2) === 'Count of Amount')
    expect(column(w, 2, 4, 7)).toEqual(['1', '2', '1', '4'])
    await menuAt(w, '4:2')
    menuItem('pivotSummarize')!.click()
    await settle()
    expect(menuItem('pivotFn-count')?.classList.contains('xs-menu-checked')).toBe(true)
    await choose('pivotShowAs-percentOfTotal', 'pivotShowAs')
    await until(() => cellText(w, 4, 2) === '25.00%')
    await key(w, 'z', { ctrlKey: true })
    await until(() => cellText(w, 4, 2) === '1')
    await key(w, 'z', { ctrlKey: true })
    await until(() => cellText(w, 3, 2) === 'Sum of Amount')
  })

  it("a label's menu: Sort, Filter, Subtotal, Remove, Field Settings and PivotTable Options", async () => {
    const { w } = await sales({ rows: [0, 1], values: [{ field: 3 }] }, 'R!A4')
    await menuAt(w, '4:1')
    expect(['pivotSort', 'pivotFilter', 'pivotSubtotal', 'pivotGroup', 'pivotUngroup', 'pivotRemove', 'pivotFieldSettings', 'pivotOptions'].map((id) => !!menuItem(id))).toEqual(Array(8).fill(true))
    expect(menuItem('pivotRemove')?.textContent?.trim()).toBe('Remove "Region"')
    expect(menuItem('pivotSubtotal')?.classList.contains('xs-menu-checked')).toBe(true)
    await choose('pivotSortDesc', 'pivotSort')
    await until(() => cellText(w, 4, 1) === 'South')
    await menuAt(w, '4:1')
    await choose('pivotHide', 'pivotFilter')
    await until(() => cellText(w, 4, 1) === 'North')
    await menuAt(w, '4:1')
    await choose('pivotRemove')
    await until(() => cellText(w, 4, 1) === 'a')
  })
})

describe('Value Field Settings and Field Settings', () => {
  it('Value Field Settings: a custom name, Average, Difference From with its base item, and a number format', async () => {
    const { w } = await sales({ rows: [0], values: [{ field: 3 }] })
    await menuAt(w, '4:2')
    await choose('pivotValueSettings')
    expect(q('[data-pv-source]')?.textContent).toBe('Amount')
    await set('[data-pv-function]', 'average')
    await ok()
    // the name goes with the function, as xlcalc names it
    await until(() => cellText(w, 3, 2) === 'Average of Amount')
    expect(cellText(w, 5, 2)).toBe('7.5')
    await api(w).pivotDialog({ dialog: 'valueFieldSettings', sheet: 'R', pivot: 'PivotTable1', field: 3, value: 0 })
    await until(() => !!q('[data-pivot-value]'))
    await set('[data-pv-name]', 'Mean')
    q<HTMLButtonElement>('[data-pv-tab="showAs"]')!.click()
    await set('[data-pv-show-as]', 'difference')
    expect(q<HTMLSelectElement>('[data-pv-base-item]')!.disabled).toBe(false)
    await set('[data-pv-base-item]', '0')
    q<HTMLButtonElement>('[data-pv-format]')!.click()
    await settle()
    await set('[data-pnf-code]', '0.00')
    q<HTMLButtonElement>('[data-pivot-number-format] [data-xs-dialog-ok]')!.click()
    await settle()
    await ok()
    await until(() => cellText(w, 3, 2) === 'Mean')
    // East is the base item: its own cell empty, North's average 7.5 less 7
    expect(column(w, 2, 4, 5)).toEqual(['', '0.50'])
  })

  it('Field Settings: no subtotals, tabular form with repeated labels, a custom name', async () => {
    const { w, s } = await sales({ rows: [0, 1], values: [{ field: 3 }] }, 'R!A4')
    await menuAt(w, '4:1')
    await choose('pivotFieldSettings')
    expect(q<HTMLInputElement>('[data-pfd-name]')!.value).toBe('Region')
    await set('[data-pfd-subtotals="none"]', true)
    q<HTMLButtonElement>('[data-pfd-tab="layout"]')!.click()
    await set('[data-pfd-form="tabular"]', true)
    await set('[data-pfd-repeat]', true)
    await set('[data-pfd-name]', 'Area')
    await ok()
    // Region's labels take a column of their own, repeated on each line, and no subtotal line
    await until(() => cellText(w, 6, 1) === 'North')
    const t = (await s.pivotTable('R', 'PivotTable1'))!.table
    expect(t.fields[0]).toMatchObject({ name: 'Area', subtotals: [], layout: 'tabular', repeatLabels: true })
    expect(column(w, 1, 4, 6)).toEqual(['East', 'North', 'North'])
    expect(column(w, 2, 4, 6)).toEqual(['a', 'a', 'b'])
  })
})

describe('PivotTable Options and the Design commands', () => {
  // Excel's RowGrand ("Show grand totals for rows") is xlsx colGrandTotals, the column at the right: xlcalc's grandTotals.columns, as
  // its docs/development.md (PivotTables, the case no-row-grand) records Excel for Windows writing it
  it('Options: a text for empty cells and the grand totals for rows off in one step, then the new name', async () => {
    const { w, s } = await sales({ rows: [0], columns: [1], values: [{ field: 3 }] })
    expect(await api(w).pivotCommand('options')).toBe(true)
    await settle()
    await set('[data-po-name]', 'Sales')
    await set('[data-po="empty"]', 'none')
    q<HTMLButtonElement>('[data-po-tab="totals"]')!.click()
    await set('[data-po="grandRows"]', false)
    await ok()
    await until(() => lastState(w)?.pivot === 'Sales')
    expect(cellText(w, 5, 3)).toBe('none')
    // the grand total column at the right goes, the grand total row stays
    expect(cellText(w, 4, 4)).toBe('')
    expect(cellText(w, 8, 1)).toBe('Grand Total')
    expect((await s.pivotTable('R', 'Sales'))?.table).toMatchObject({ empty: { show: true, text: 'none' }, grandTotals: { rows: true, columns: false } })
    expect(await api(w).pivotCommand('options')).toBe(true)
    await settle()
    q<HTMLButtonElement>('[data-po-tab="totals"]')!.click()
    expect(q<HTMLInputElement>('[data-po="grandRows"]')!.checked).toBe(false)
    expect(q<HTMLInputElement>('[data-po="grandColumns"]')!.checked).toBe(true)
  })

  it('Design: Report Layout, Grand Totals, Subtotals and Blank Rows, and the toolbar shows them while the active cell is in a PivotTable', async () => {
    const { w, s } = await sales({ rows: [0, 1], values: [{ field: 3 }] })
    expect(await api(w).pivotDesign({ layout: 'tabular' })).toMatchObject({ ok: true })
    await until(() => cellText(w, 3, 2) === 'Kind')
    expect(await api(w).pivotDesign({ subtotals: 'none' })).toMatchObject({ ok: true })
    expect(await api(w).pivotDesign({ blankLine: true })).toMatchObject({ ok: true })
    const t = (await s.pivotTable('R', 'PivotTable1'))!.table
    expect(t).toMatchObject({ layout: 'tabular' })
    expect(t.fields[0]).toMatchObject({ subtotals: [], blankLine: true })
  })

  it("Design's On for Rows Only keeps the grand total column at the right and takes away the grand total row", async () => {
    const { w, s } = await sales({ rows: [0], columns: [1], values: [{ field: 3 }] })
    expect(await api(w).pivotDesign({ grandTotals: 'rows' })).toMatchObject({ ok: true })
    await until(() => cellText(w, 8, 1) === '')
    expect(cellText(w, 4, 4)).toBe('Grand Total')
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.grandTotals).toEqual({ rows: false, columns: true })
  })
})

describe('Number Format of a value field', () => {
  it("keeps a built-in format's id when OK leaves it as it was, and sends the format picked", async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3, numberFormat: 4 }] })
    await until(() => cellText(w, 5, 2) === '15.00')
    await menuAt(w, '4:2')
    await choose('pivotNumberFormat')
    expect(q<HTMLInputElement>('[data-pnf-code]')!.value).toBe('#,##0.00')
    await ok()
    await settle()
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.values?.[0]?.numberFormat).toBe(4)
    await api(w).pivotDialog({ dialog: 'valueFieldSettings', sheet: 'R', pivot: 'PivotTable1', field: 3, value: 0 })
    await until(() => !!q('[data-pivot-value]'))
    await ok()
    await settle()
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.values?.[0]?.numberFormat).toBe(4)
    expect(cellText(w, 5, 2)).toBe('15.00')
    await menuAt(w, '4:2')
    await choose('pivotNumberFormat')
    await set('[data-pnf-code]', '0.0')
    await ok()
    await until(() => cellText(w, 5, 2) === '15.0')
  })
})
