// Grouping, Ungroup and Insert Calculated Field and Item on SheetView over localSource: dates by periods and numbers by an
// interval in the Grouping dialog, items selected grouped by hand, Ungroup, a calculated field put on Values as Excel's OK
// does, a calculated item, Delete; each through xlcalc's operation.
import { mount } from '@vue/test-utils'
import { describe, expect, it } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import { SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '../src'
import { localSource } from '../src/local'
import { cellText, goTo, settle, until } from './helpers'
import { addSales, salesBook } from './pivot-fixture'
import { api, choose, menuAt, ok, q, sales, set } from './pivot-sheet'

describe('Grouping and Ungroup', () => {
  it('groups dates by months and years from the cell menu, and Ungroup takes the periods away', async () => {
    const { w, s } = await sales({ rows: [4], values: [{ field: 3 }] }, 'R!A4')
    await menuAt(w, '4:1')
    await choose('pivotGroup')
    await until(() => !!q('[data-pivot-group]'))
    expect(q<HTMLInputElement>('[data-pg-start]')!.value).toBe('4/1/2024')
    // a field of whole days ends a day after its last date, as Excel proposes it
    expect(q<HTMLInputElement>('[data-pg-end]')!.value).toBe('3/10/2025')
    expect([...q<HTMLSelectElement>('[data-pg-periods]')!.selectedOptions].map((o) => o.value)).toEqual(['months'])
    await set('[data-pg-periods]', ['months', 'years'])
    expect(q<HTMLInputElement>('[data-pg-days]')!.disabled).toBe(true)
    await ok()
    await until(() => cellText(w, 4, 1) === '2024')
    const cache = (await s.pivotTable('R', 'PivotTable1'))!.cache
    expect(cache.fields.map((f) => f.name)).toContain('Years (Date)')
    await goTo(w, 'R!A4')
    expect(await api(w).ungroupPivot()).toMatchObject({ ok: true })
    await until(() => cellText(w, 4, 1) === '4/1/2024' || cellText(w, 4, 1)!.startsWith('#'))
    expect((await s.pivotTable('R', 'PivotTable1'))!.cache.fields.map((f) => f.name)).not.toContain('Years (Date)')
  })

  it('groups numbers by an interval with a start of its own, and refuses a box that is not a number', async () => {
    const { w } = await sales({ rows: [3], values: [{ field: 3, function: 'count' }] }, 'R!A4')
    expect(await api(w).pivotCommand('groupField')).toBe(true)
    await settle()
    expect(q<HTMLInputElement>('[data-pg-by]')!.value).toBe('10')
    await set('[data-pg-start]', 'x')
    await ok()
    expect(q('[data-pg-error]')?.textContent).toBe('Enter a number.')
    expect(q<HTMLInputElement>('[data-pg-auto-start]')!.checked).toBe(false)
    await set('[data-pg-start]', '0')
    await ok()
    await until(() => cellText(w, 4, 1) === '0-9')
    // the last group of whole numbers takes the end in, as xlcalc's docs/pivot-tables.md (Grouping) says Excel does
    expect(cellText(w, 5, 1)).toBe('10-20')
  })

  it('groups the items selected by hand into Group1, and Ungroup the group selected', async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!A4')
    await goTo(w, 'R!A4:A5')
    expect(await api(w).groupPivot()).toBe(true)
    await until(() => cellText(w, 4, 1) === 'Group1')
    expect((await s.pivotTable('R', 'PivotTable1'))!.table.rows).toEqual([5, 0])
    await goTo(w, 'R!A4')
    expect(await api(w).ungroupPivot()).toMatchObject({ ok: true })
    await until(() => cellText(w, 4, 1) !== 'Group1')
  })
})

describe('Insert Calculated Field and Calculated Item', () => {
  it('adds a calculated field named Field1 and puts it on Values, as Excel does', async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] })
    expect(await api(w).pivotCommand('calculatedField')).toBe(true)
    await settle()
    expect(q<HTMLInputElement>('[data-pc-name]')!.value).toBe('Field1')
    expect(q<HTMLInputElement>('[data-pc-formula]')!.value).toBe('= 0')
    await set('[data-pc-name]', 'Double')
    await set('[data-pc-formula]', '=')
    await set('[data-pc-fields]', '3')
    q<HTMLButtonElement>('[data-pc-insert-field]')!.click()
    await settle()
    expect(q<HTMLInputElement>('[data-pc-formula]')!.value).toBe('=Amount')
    await set('[data-pc-formula]', '=Amount*2')
    await ok()
    await until(() => cellText(w, 3, 3) === 'Sum of Double')
    expect(cellText(w, 4, 3)).toBe('14')
    const got = (await s.pivotTable('R', 'PivotTable1'))!
    expect(got.cache.fields.at(-1)).toMatchObject({ name: 'Double', formula: 'Amount*2' })
    expect(got.table.values).toMatchObject([{ field: 3 }, { field: 5, name: 'Sum of Double' }])
    // one step of undo takes the field and its place on Values away
    expect(await s.undo()).toMatchObject({ ok: true })
    const back = (await s.pivotTable('R', 'PivotTable1'))!
    expect(back.cache.fields.map((f) => f.name)).not.toContain('Double')
    expect(back.table.values).toMatchObject([{ field: 3 }])
  })

  it('adds a calculated item of the field under the active cell, and Delete removes it', async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!A4')
    expect(await api(w).pivotCommand('calculatedItem')).toBe(true)
    await settle()
    expect(q('[data-pivot-calc] .xs-modal-title')?.textContent).toBe('Insert Calculated Item in "Region"')
    expect(q<HTMLInputElement>('[data-pc-name]')!.value).toBe('Formula1')
    await set('[data-pc-formula]', '=')
    await set('[data-pc-items]', '1')
    q<HTMLButtonElement>('[data-pc-insert-item]')!.click()
    await settle()
    expect(q<HTMLInputElement>('[data-pc-formula]')!.value).toBe('=North')
    await set('[data-pc-formula]', '=North+South')
    await set('[data-pc-name]', 'Both')
    await ok()
    await until(() => cellText(w, 7, 1) === 'Both')
    expect(cellText(w, 7, 2)).toBe('35')
    expect(await api(w).pivotCommand('calculatedItem')).toBe(true)
    await settle()
    await set('[data-pc-name]', 'Both')
    expect(q<HTMLInputElement>('[data-pc-formula]')!.value).toBe('=North+South')
    expect(q('[data-pc-add]')?.textContent).toBe('Modify')
    q<HTMLButtonElement>('[data-pc-delete]')!.click()
    await until(() => cellText(w, 7, 1) === 'Grand Total')
    expect((await s.pivotTable('R', 'PivotTable1'))!.cache.calculatedItems ?? []).toEqual([])
  })
})

describe("the toolbar's PivotTable Analyze and Design commands", () => {
  it('show while the active cell is in a PivotTable, the Design lists checking the item in force', async () => {
    const s = localSource(salesBook(), { calc: 'inline' })
    expect(await s.structure(addSales({ rows: [0, 1], values: [{ field: 3 }] }))).toMatchObject({ ok: true })
    const sheet = ref<SheetViewApi | null>(null)
    const state = ref<SheetViewState | null>(null)
    const Host = defineComponent(() => () => [h(SheetToolbar, { source: s, sheet: sheet.value, state: state.value }), h(SheetView, { ref: sheet, source: s, style: 'height: 400px', pivotFields: false, onState: (x: SheetViewState | null) => (state.value = x) })])
    const w = mount(Host, { attachTo: document.body })
    await until(() => !!sheet.value && !!state.value)
    const view = w.findComponent(SheetView)
    await goTo(view, 'R!A1')
    expect(w.find('[data-tb="pivot-subtotals"]').exists()).toBe(false)
    await goTo(view, 'R!B4')
    await until(() => state.value?.canEditPivot === true)
    await w.find('[data-tb="pivot-layout"]').trigger('click')
    await until(() => w.find('[data-tb="pivot-layout-compact"]').attributes('aria-checked') === 'true')
    await w.find('[data-tb="pivot-layout-tabular"]').trigger('click')
    await until(() => cellText(view, 3, 2) === 'Kind')
    await w.find('[data-tb="pivot-grandTotals"]').trigger('click')
    await until(() => w.find('[data-tb="pivot-grandTotals-on"]').attributes('aria-checked') === 'true')
    await w.find('[data-tb="pivot-calculations"]').trigger('click')
    await w.find('[data-tb="pivot-calc-field"]').trigger('click')
    await until(() => !!q('[data-pivot-calc]'))
    w.unmount()
    await s.close()
  })
})
