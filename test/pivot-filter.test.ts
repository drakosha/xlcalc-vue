// The filter menu of a PivotTable's field on SheetView over localSource: its buttons on Row Labels and a Filters area's item,
// Alt+Down, the items' list with its search, Label, Value and Date Filters with their dialogs, Top 10, Clear Filter From,
// the sorts and More Sort Options, each one editPivotTable.
import { describe, expect, it } from 'vitest'
import { cellText, key, settle, until } from './helpers'
import { ok, q, sales, set } from './pivot-sheet'

const menu = () => q('[data-pivot-filter-menu]')
async function button(cell: string): Promise<void> {
  await until(() => !!q(`[data-pivot-button="${cell}"]`))
  q<HTMLButtonElement>(`[data-pivot-button="${cell}"]`)!.click()
  await settle()
}
async function pick(sel: string): Promise<void> {
  q<HTMLButtonElement>(sel)!.click()
  await settle()
}
const listed = () => [...document.querySelectorAll('[data-pfm-values] .xs-value-text')].map((x) => x.textContent)

describe('the filter menu of a PivotTable', () => {
  it("unchecks items on Row Labels' list, searches it, and clears the filter", async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] })
    await button('A3')
    expect(listed()).toEqual(['(Select All)', 'East', 'North', 'South'])
    await set('[data-pfm-item="1"]', false)
    await pick('[data-pfm-ok]')
    await until(() => cellText(w, 5, 1) === 'South')
    expect(q('[data-pivot-button="A3"]')?.classList.contains('xs-on')).toBe(true)
    await button('A3')
    await set('[data-pfm-search]', 'th')
    expect(listed()).toEqual(['(Select All)', 'North', 'South'])
    await set('[data-pfm-item="1"]', true)
    await pick('[data-pfm-ok]')
    // the search leaves East out of what OK shows
    await until(() => cellText(w, 4, 1) === 'North')
    expect(cellText(w, 5, 1)).toBe('South')
    await button('A3')
    await pick('[data-pfm="clear"]')
    await until(() => cellText(w, 4, 1) === 'East')
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.fields[0]?.items?.some((x) => x.hidden)).toBe(false)
  })

  it('Label Filters: Begins With and Between; Value Filters: Greater Than and Top 10', async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] })
    await button('A3')
    await pick('[data-pfm="label"]')
    await pick('[data-pfm-op="beginsWith"]')
    await until(() => q('[data-pivot-filter-dialog] .xs-modal-title')?.textContent === 'Label Filter (Region)')
    await set('[data-pf-first]', 'n*')
    await ok()
    await until(() => cellText(w, 4, 1) === 'North')
    expect(cellText(w, 5, 1)).toBe('Grand Total')
    await button('A3')
    await pick('[data-pfm="value"]')
    await pick('[data-pfm-op="greaterThan"]')
    await set('[data-pf-first]', '10')
    await ok()
    // a value filter replaces the label filter without Allow multiple filters per field
    await until(() => cellText(w, 6, 1) === 'Grand Total')
    expect([cellText(w, 4, 1), cellText(w, 5, 1)]).toEqual(['North', 'South'])
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.filters).toEqual([{ field: 0, type: 'valueGreaterThan', args: [10], value: 0 }])
    await button('A3')
    await pick('[data-pfm="value"]')
    await pick('[data-pfm-op="top10"]')
    await set('[data-pf-count]', '1')
    await ok()
    await until(() => cellText(w, 5, 1) === 'Grand Total')
    expect(cellText(w, 4, 1)).toBe('South')
    await button('A3')
    await pick('[data-pfm="label"]')
    await pick('[data-pfm-op="between"]')
    await set('[data-pf-first]', 'a')
    await set('[data-pf-second]', 'o')
    await ok()
    await until(() => cellText(w, 6, 1) === 'Grand Total')
    expect([cellText(w, 4, 1), cellText(w, 5, 1)]).toEqual(['East', 'North'])
  })

  it('Date Filters on a field of dates, with a date typed in the workbook locale', async () => {
    const { w, s } = await sales({ rows: [4], values: [{ field: 3 }] })
    await button('A3')
    expect(q('[data-pfm="date"]')).not.toBeNull()
    await pick('[data-pfm="date"]')
    await pick('[data-pfm-op="after"]')
    await set('[data-pf-first]', '1/1/2025')
    await ok()
    await until(() => cellText(w, 6, 1) === 'Grand Total')
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.filters).toEqual([{ field: 4, type: 'dateNewerThan', args: ['2025-01-01T00:00:00'] }])
    await button('A3')
    await pick('[data-pfm="date"]')
    await pick('[data-pfm-op="allDates"]')
    await pick('[data-pfm-op="M5"]')
    await until(() => cellText(w, 5, 1) === 'Grand Total')
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.filters).toEqual([{ field: 4, type: 'M5' }])
  })

  it('sorts Z to A, and More Sort Options sorts by a value field', async () => {
    const { w, s } = await sales({ rows: [0], values: [{ field: 3 }] })
    await button('A3')
    await pick('[data-pfm="sortDesc"]')
    await until(() => cellText(w, 4, 1) === 'South')
    await button('A3')
    await pick('[data-pfm="moreSort"]')
    expect(q<HTMLInputElement>('[data-pf-sort="descending"]')!.checked).toBe(true)
    await set('[data-pf-sort="ascending"]', true)
    await set('[data-pf-asc-by]', '0')
    expect(q('[data-pf-summary]')?.textContent).toBe('Sort Sum of Amount in ascending order')
    await ok()
    await until(() => cellText(w, 4, 1) === 'East')
    expect([cellText(w, 5, 1), cellText(w, 6, 1)]).toEqual(['North', 'South'])
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.fields[0]?.sort).toEqual({ order: 'ascending', by: { value: 0 } })
  })

  it("a Filters area's item: Alt+Down opens its list, one item or several shown", async () => {
    const { w } = await sales({ rows: [0], pages: [{ field: 1 }], values: [{ field: 3 }] }, 'R!B1')
    await until(() => !!q('[data-pivot-button="B1"]'))
    await key(w, 'ArrowDown', { altKey: true })
    await settle()
    expect(menu()).not.toBeNull()
    await set('[data-pfm-page="1"]', true)
    await pick('[data-pfm-ok]')
    await until(() => cellText(w, 1, 2) === 'b')
    expect([cellText(w, 4, 1), cellText(w, 5, 1)]).toEqual(['North', 'South'])
    await button('B1')
    await set('[data-pfm-multiple]', true)
    expect(listed()[0]).toBe('(Select All)')
    await pick('[data-pfm-ok]')
    await until(() => cellText(w, 1, 2) === '(All)')
  })

  it("opens no field's menu, and shows no funnel, for a button that filters no field: Σ Values alone on the rows, Σ Values' name", async () => {
    const { s } = await sales({ columns: [0], rows: ['values'], values: [{ field: 3 }, { field: 2 }] }, 'R!B5')
    // Region on the columns filtered: its own button lights, Row Labels over Σ Values does not
    expect(await s.structure({ op: 'editPivotTable', sheet: 'R', pivot: 'PivotTable1', set: { fields: [{ field: 0, items: [{ item: 0, hidden: true }, { item: 1 }, { item: 2 }] }] } })).toMatchObject({ ok: true })
    await until(() => !!q('[data-pivot-button="B3"]')?.classList.contains('xs-on'))
    await button('A4')
    expect(menu()).toBeNull()
    expect(q('[data-pivot-button="A4"]')?.classList.contains('xs-on')).toBe(false)
    document.body.innerHTML = ''
    await sales({ rows: [0], columns: ['values', 1], values: [{ field: 3 }, { field: 2 }], layout: 'outline' }, 'R!B6')
    await button('B3')
    expect(menu()).toBeNull()
  })

  it("a compact row field's name in its own column has its own button and funnel; Row Labels' funnel is for the fields without one", async () => {
    const { s } = await sales({ rows: [0, 1], values: [{ field: 3 }], fields: [{ field: 0, layout: 'tabular' }] }, 'R!C5')
    // Kind (field 1) filtered: its own button B3 lights, Row Labels A3 does not
    expect(await s.structure({ op: 'editPivotTable', sheet: 'R', pivot: 'PivotTable1', set: { fields: [{ field: 1, items: [{ item: 0, hidden: true }, { item: 1 }] }] } })).toMatchObject({ ok: true })
    await until(() => !!q('[data-pivot-button="B3"]')?.classList.contains('xs-on'))
    expect(q('[data-pivot-button="A3"]')?.classList.contains('xs-on')).toBe(false)
    await button('B3')
    expect(listed()).toEqual(['(Select All)', 'a', 'b'])
  })

  it("Select field picks among the row fields of Row Labels in compact form, and Esc closes the menu", async () => {
    const { w } = await sales({ rows: [0, 1], values: [{ field: 3 }] })
    await button('A3')
    await set('[data-pfm-field]', '1')
    expect(listed()).toEqual(['(Select All)', 'a', 'b'])
    menu()!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    await settle()
    expect(menu()).toBeNull()
    void w
  })
})
