// The rules the dialogs over a PivotTable follow, as Excel for Windows does them: Custom Name after another function,
// Σ Values kept on an axis when value fields come and go, Ungroup where xlcalc ungroups, a Filters area's funnel, Number of
// days checked, the base item reset with its base field, formulas in the workbook's locale, the commands gated by
// canEditPivot, and a request the sheet cannot open passed on to the host.
import { describe, expect, it } from 'vitest'
import type { SheetSource } from '../src'
import { cellText, goTo, key, settle, until } from './helpers'
import { api, choose, lastState, menuAt, menuItem, ok, q, sales, set } from './pivot-sheet'

const openValue = async (w: Parameters<typeof api>[0]) => {
  await api(w).pivotDialog({ dialog: 'valueFieldSettings', sheet: 'R', pivot: 'PivotTable1', field: 3, value: 0 })
  await until(() => !!q('[data-pivot-value]'))
}
const nameBox = () => q<HTMLInputElement>('[data-pv-name]')!

describe('Value Field Settings: Custom Name follows the function picked, as Excel for Windows does', () => {
  it("replaces the old function's prefix in any case and keeps the rest; resets another name; a name typed after the pick stays", async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] })
    await openValue(w)
    await set('[data-pv-function]', 'count')
    expect(nameBox().value).toBe('Count of Amount')
    await set('[data-pv-name]', 'sum of amount2')
    await set('[data-pv-function]', 'sum')
    // the prefix is Count's now: another name is reset
    expect(nameBox().value).toBe('Sum of Amount')
    await set('[data-pv-name]', 'count of amount2')
    // case 4: a name typed before the pick is overwritten; 4b: one typed after it is kept
    await set('[data-pv-name]', 'Total')
    await set('[data-pv-function]', 'max')
    expect(nameBox().value).toBe('Max of Amount')
    await set('[data-pv-name]', 'Total')
    await ok()
    await until(() => cellText(w, 3, 2) === 'Total')
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.values?.[0]).toMatchObject({ name: 'Total', function: 'max' })
    // Summarize Values By in the cell menu: a name without the prefix becomes the new function's
    await menuAt(w, '4:2')
    await choose('pivotFn-count', 'pivotSummarize')
    await until(() => cellText(w, 3, 2) === 'Count of Amount')
  })

  it('keeps the rest of a prefixed name as typed, in ru-RU too', async () => {
    const { w } = await sales({ rows: [0], values: [{ field: 3, name: 'Сумма по полю Amount2' }] }, 'R!B4', 'ru-RU')
    await openValue(w)
    await set('[data-pv-function]', 'count')
    expect(nameBox().value).toBe('Количество по полю Amount2')
  })

  it("Summarize Values By leaves the name to xlcalc's editPivotTable, which keeps the rest of a prefixed name", async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3, name: 'Sum of Amount2' }] })
    await menuAt(w, '4:2')
    await choose('pivotFn-count', 'pivotSummarize')
    await until(() => cellText(w, 3, 2) === 'Count of Amount2')
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.values?.[0]).toMatchObject({ name: 'Count of Amount2', function: 'count' })
  })

  it('resets the base item to (previous) when another base field is picked', async () => {
    const { w } = await sales({ rows: [0], columns: [1], values: [{ field: 3, showAs: 'difference', baseField: 0, baseItem: 1 }] })
    await openValue(w)
    expect(q<HTMLSelectElement>('[data-pv-base-item]')!.value).toBe('1')
    await set('[data-pv-base-field]', '1')
    expect(q<HTMLSelectElement>('[data-pv-base-item]')!.value).toBe('previous')
  })
})

describe('Σ Values with the value fields', () => {
  it('Remove of a value field takes Σ Values off its axis when one value field is left', async () => {
    const { w, s } = await sales({ rows: [0], columns: ['values'], values: [{ field: 3 }, { field: 3, function: 'count' }] })
    await menuAt(w, '4:3')
    expect(menuItem('pivotRemoveValue')?.textContent?.trim()).toBe('Remove "Count of Amount"')
    await choose('pivotRemoveValue')
    await until(() => cellText(w, 3, 3) === '')
    const t = (await s.pivotTable('R', 'PivotTable1'))!.table
    expect(t.values?.length).toBe(1)
    expect(t.columns ?? []).not.toContain('values')
  })

  it("OK in Insert Calculated Field puts the field second with Σ Values on Columns, or alone without Σ, as Excel's OK", async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] })
    expect(await api(w).pivotCommand('calculatedField')).toBe(true)
    await settle()
    await set('[data-pc-formula]', '=Amount*2')
    await ok()
    await until(() => cellText(w, 3, 3) === 'Sum of Field1')
    const t = (await s.pivotTable('R', 'PivotTable1'))!.table
    expect(t.values?.map((v) => v.field)).toEqual([3, 5])
    expect(t.columns).toEqual(['values'])
    const alone = await sales({ rows: [0] }, 'R!A4')
    expect(await api(alone.w).pivotCommand('calculatedField')).toBe(true)
    await settle()
    await set('[data-pc-formula]', '=Amount*2')
    await ok()
    await until(() => cellText(alone.w, 3, 2) === 'Sum of Field1')
    const one = (await alone.s.pivotTable('R', 'PivotTable1'))!.table
    expect(one.values?.length).toBe(1)
    expect(one.columns ?? []).toEqual([])
  })
})

describe('formulas of calculated fields in the workbook locale', () => {
  it('reads a ru-RU formula as typed there and shows an existing one in it', async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!B4', 'ru-RU')
    expect(await api(w).pivotCommand('calculatedField')).toBe(true)
    await settle()
    await set('[data-pc-name]', 'Half')
    await set('[data-pc-formula]', '=ЕСЛИ(Amount>0;Amount*0,5;0)')
    await ok()
    await until(() => cellText(w, 4, 3) === '3,5')
    expect((await s.pivotTable('R', 'PivotTable1'))!.cache.fields.at(-1)).toMatchObject({ name: 'Half', formula: 'IF(Amount>0,Amount*0.5,0)' })
    expect(await api(w).pivotCommand('calculatedField')).toBe(true)
    await settle()
    await set('[data-pc-name]', 'Half')
    expect(q<HTMLInputElement>('[data-pc-formula]')!.value).toBe('=ЕСЛИ(Amount>0;Amount*0,5;0)')
  })

  it("lists a calculated item's formula of a field whose own name holds a character of a pattern", async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!A4')
    expect(await s.structure({ op: 'setPivotCalculatedItem', sheet: 'R', pivot: 'PivotTable1', field: 0, name: 'Both', formula: '=North+South' })).toMatchObject({ ok: true })
    expect(await s.structure({ op: 'editPivotTable', sheet: 'R', pivot: 'PivotTable1', set: { fields: [{ field: 0, name: 'Reg*' }] } })).toMatchObject({ ok: true })
    await goTo(w, 'R!A4')
    expect(await api(w).pivotCommand('calculatedItem')).toBe(true)
    await settle()
    await set('[data-pc-name]', 'Both')
    expect(q<HTMLInputElement>('[data-pc-formula]')!.value).toBe('=North+South')
  })
})

describe('Grouping and Ungroup', () => {
  it('offers Ungroup on a group of items grouped by hand, not on an item of its base field', async () => {
    const { w } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!A4')
    await goTo(w, 'R!A4:A5')
    expect(await api(w).groupPivot()).toBe(true)
    await until(() => cellText(w, 4, 1) === 'Group1')
    // Group1 at A4 over East (A5) and North (A6), South its own group
    await goTo(w, 'R!A5')
    await menuAt(w, '5:1')
    await until(() => menuItem('pivotUngroup')?.disabled === true)
    await menuAt(w, '4:1')
    await until(() => menuItem('pivotUngroup')?.disabled === false)
  })

  it('refuses Number of days that is not a whole number, without grouping', async () => {
    const { w, s } = await sales({ rows: [4], values: [{ field: 3 }] }, 'R!A4')
    expect(await api(w).pivotCommand('groupField')).toBe(true)
    await settle()
    await set('[data-pg-periods]', ['days'])
    await set('[data-pg-days]', 'x')
    await ok()
    expect(q('[data-pg-error]')?.textContent).toBe('Enter a number.')
    expect((await s.pivotTable('R', 'PivotTable1'))!.cache.fields[4]?.group).toBeUndefined()
  })
})

describe('the filter buttons and the commands, gated by canEditPivot', () => {
  it("Row Labels, which xlcalc names no field, has the menu and the commands of the rows' first field, as before (not measured in Excel)", async () => {
    const { w } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!A3')
    await menuAt(w, '3:1')
    expect(['pivotSort', 'pivotFilter', 'pivotRemove', 'pivotFieldSettings', 'pivotOptions'].map((id) => !!menuItem(id))).toEqual([true, true, true, true, true])
    expect(menuItem('pivotRemove')?.textContent?.trim()).toBe('Remove "Region"')
    expect(await api(w).pivotCommand('calculatedItem')).toBe(true)
    await settle()
    expect(q('[data-pivot-calc]')).not.toBeNull()
    document.body.innerHTML = ''
    const dates = await sales({ rows: [4], values: [{ field: 3 }] }, 'R!A3')
    expect(await api(dates.w).pivotCommand('groupField')).toBe(true)
  })

  it("a Filters area's button shows the funnel with one item chosen", async () => {
    const { w, s } = await sales({ rows: [0], pages: [{ field: 1 }], values: [{ field: 3 }] }, 'R!B4')
    await until(() => !!q('[data-pivot-button="B1"]'))
    expect(q('[data-pivot-button="B1"]')!.classList.contains('xs-on')).toBe(false)
    expect(await api(w).structure({ op: 'editPivotTable', sheet: 'R', pivot: 'PivotTable1', set: { pages: [{ field: 1, item: 0 }] } })).toMatchObject({ ok: true })
    await until(() => cellText(w, 1, 2) === 'a')
    await until(() => !!q('[data-pivot-button="B1"]')?.classList.contains('xs-on'))
    void s
  })

  it('a protected sheet without Use PivotTable and PivotChart has no buttons, no Alt+Down menu, no Group or Ungroup', async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!A3')
    await until(() => !!q('[data-pivot-button="A3"]'))
    expect(await s.protect('R', {})).toMatchObject({ ok: true })
    await until(() => !!lastState(w)?.protection)
    await settle()
    expect(q('[data-pivot-button]')).toBeNull()
    await key(w, 'ArrowDown', { altKey: true })
    await settle()
    expect(q('[data-pivot-filter-menu]')).toBeNull()
    expect(await api(w).groupPivot()).toBe(false)
    expect(await api(w).ungroupPivot()).toBeNull()
  })

  it("passes the pane's request on to the host when the source cannot tell what a PivotTable's cells are", async () => {
    const { w } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!B4', undefined, { wrap: (s) => ({ ...s, pivotCells: undefined }) as SheetSource, editable: false })
    expect(api(w).pivotDialog({ dialog: 'fieldSettings', sheet: 'R', pivot: 'PivotTable1', field: 0 })).toBe(true)
    await settle()
    expect(w.emitted('pivotDialog')?.at(-1)?.[0]).toMatchObject({ dialog: 'fieldSettings', field: 0 })
    expect(q('[data-pivot-field]')).toBeNull()
  })
})
