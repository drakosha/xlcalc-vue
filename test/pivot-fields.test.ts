// The PivotTable Fields pane and Insert PivotTable on SheetView over localSource: the pane follows the active cell and the
// Field List, a checked field goes where Excel puts it, the field menus move fields, Defer Layout Update waits for Update,
// each change is one undo step, the dialogs a menu opens are the sheet's own, and Insert PivotTable creates one on a new sheet or a place.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { SheetViewApi, SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { cellText, goTo, mountSheet, settle, until } from './helpers'
import { addSales, pivotBook, salesBook } from './pivot-fixture'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

async function report(book = pivotBook(), props: Record<string, unknown> = {}): Promise<{ w: VueWrapper; s: LocalSource }> {
  const s = localSource(book, { calc: 'inline' })
  sources.push(s)
  const w = await mountSheet(s, { style: 'height: 400px', ...props })
  await goTo(w, 'Report!B4')
  await until(() => !!pane())
  await until(() => chips('rows').length === 1)
  return { w, s }
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
const pane = () => document.querySelector<HTMLElement>('[data-xs-pf]')
const chips = (area: string) => [...document.querySelectorAll<HTMLElement>(`[data-xs-pf-area="${area}"] [data-xs-pf-chip]`)].map((x) => x.textContent?.trim())
const box = (field: number) => document.querySelector<HTMLInputElement>(`[data-xs-pf-field="${field}"] input[type="checkbox"]`)!
const item = (id: string) => document.querySelector<HTMLButtonElement>(`[data-xs-pf-item="${id}"]`)
async function chipMenu(area: string, i: number): Promise<void> {
  document.querySelectorAll<HTMLButtonElement>(`[data-xs-pf-area="${area}"] .xs-pf-chip`)[i]!.click()
  await settle()
}
async function setBox(field: number, on: boolean): Promise<void> {
  const b = box(field)
  b.checked = on
  b.dispatchEvent(new Event('change'))
  await settle()
}

describe('the PivotTable Fields pane', () => {
  it('shows while the active cell is in a PivotTable and the Field List is on, with the fields and the areas', async () => {
    const { w } = await report()
    expect(lastState(w)?.fieldList).toBe(true)
    expect([...document.querySelectorAll('[data-xs-pf-field] .xs-pf-name')].map((x) => x.textContent)).toEqual(['Region', 'Amount', 'Kind', 'Year'])
    expect([0, 1, 2, 3].map((i) => box(i).checked)).toEqual([true, true, false, false])
    expect(chips('values')).toEqual(['Sum of Amount'])
    await goTo(w, 'D10')
    expect(pane()).toBeNull()
    await goTo(w, 'A4')
    expect(pane()).not.toBeNull()
    // the close button turns the Field List off; the cell menu's Show Field List turns it on
    document.querySelector<HTMLButtonElement>('[data-xs-pf-close]')!.click()
    await settle()
    expect(pane()).toBeNull()
    expect(lastState(w)?.fieldList).toBeUndefined()
    await w.find('[data-cell="4:1"]').trigger('contextmenu', { button: 2, clientX: 40, clientY: 60 })
    await settle()
    const show = document.querySelector<HTMLButtonElement>('[data-xs-menu="fieldList"]')!
    expect(show.textContent?.trim()).toBe('Show Field List')
    show.click()
    await settle()
    expect(pane()).not.toBeNull()
    expect(api(w).fieldList()).toBe(false)
  })

  it('a checked text field goes to Rows and a number field to Values, each one step of undo; unchecked it leaves', async () => {
    const { w } = await report()
    await setBox(2, true)
    await until(() => chips('rows').length === 2)
    expect(chips('rows')).toEqual(['Region', 'Kind'])
    await setBox(3, true)
    await until(() => chips('values').length === 2)
    expect(chips('values')).toEqual(['Sum of Amount', 'Sum of Year'])
    expect(chips('columns')).toEqual(['Σ Values'])
    await api(w).undo()
    await until(() => chips('values').length === 1)
    expect(chips('rows')).toEqual(['Region', 'Kind'])
    await setBox(0, false)
    await until(() => chips('rows').length === 1)
    expect(chips('rows')).toEqual(['Kind'])
  })

  it("a checked field of dates over several years is grouped by years, quarters and months, as Excel's field list does", async () => {
    const s = localSource(salesBook(), { calc: 'inline' })
    sources.push(s)
    expect(await s.structure(addSales({ values: [{ field: 3 }] }))).toMatchObject({ ok: true })
    const w = await mountSheet(s, { style: 'height: 400px' })
    await goTo(w, 'R!A4')
    await until(() => !!pane())
    await setBox(4, true)
    await until(() => chips('rows').length > 1)
    expect(chips('rows')).toEqual(['Years (Date)', 'Quarters (Date)', 'Months (Date)', 'Date'])
  })

  it("a field's menu in an area moves it, and Value Field Settings opens the sheet's dialog", async () => {
    const { w } = await report()
    await chipMenu('rows', 0)
    expect(item('moveUp')?.disabled).toBe(true)
    expect(item('toRows')?.disabled).toBe(true)
    item('toColumns')!.click()
    await until(() => chips('columns').length === 1)
    expect(chips('rows')).toEqual([])
    await until(() => cellText(w, 3, 2) === 'Column Labels')
    await chipMenu('values', 0)
    item('valueFieldSettings')!.click()
    await settle()
    await until(() => document.querySelector<HTMLInputElement>('[data-pv-name]')?.value === 'Sum of Amount')
    expect(document.querySelector('[data-pv-source]')?.textContent).toBe('Amount')
    expect(w.emitted('pivotDialog')).toBeUndefined()
  })

  it("a field's menu in the list adds it to an area", async () => {
    await report()
    document.querySelector(`[data-xs-pf-field="2"]`)!.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }))
    await settle()
    expect(item('addRows')?.disabled).toBe(false)
    item('addFilters')!.click()
    await until(() => chips('filters').length === 1)
    expect(chips('filters')).toEqual(['Kind'])
  })

  it('Defer Layout Update keeps the changes until Update, which lays them out in one step of undo', async () => {
    const { w } = await report()
    const update = () => document.querySelector<HTMLButtonElement>('[data-xs-pf-update]')!
    const defer = document.querySelector<HTMLInputElement>('[data-xs-pf-defer]')!
    defer.checked = true
    defer.dispatchEvent(new Event('change'))
    await settle()
    expect(update().disabled).toBe(true)
    await setBox(2, true)
    await setBox(3, true)
    expect(chips('rows')).toEqual(['Region', 'Kind'])
    expect(update().disabled).toBe(false)
    expect(cellText(w, 3, 2)).toBe('Sum of Amount')
    expect(cellText(w, 5, 1)).toBe('South')
    update().click()
    await until(() => cellText(w, 3, 3) === 'Sum of Year')
    await api(w).undo()
    await until(() => chips('rows').length === 1 && chips('values').length === 1)
  })
})

describe('the PivotTable Fields pane: what it takes', () => {
  it('takes no change on a read-only sheet, nor on a protected one without Use PivotTable and PivotChart', async () => {
    await report(pivotBook(), { readOnly: true })
    expect(box(2).disabled).toBe(true)
    document.body.innerHTML = ''
    await report(pivotBook({ protection: {} }))
    expect(box(2).disabled).toBe(true)
    expect(document.querySelector<HTMLInputElement>('[data-xs-pf-defer]')!.disabled).toBe(true)
    document.body.innerHTML = ''
    await report(pivotBook({ protection: { pivotTables: true } }))
    expect(box(2).disabled).toBe(false)
  })

  it("a field's menu acts on the field where it is when an item is chosen", async () => {
    const { s } = await report()
    await setBox(2, true)
    await until(() => chips('rows').length === 2)
    await chipMenu('rows', 1)
    // Region leaves Rows from elsewhere while the menu is open: Kind is first now
    await s.structure({ op: 'editPivotTable', sheet: 'Report', pivot: 'PivotTable1', set: { rows: [2] } })
    await until(() => chips('rows').length === 1)
    item('remove')!.click()
    await until(() => chips('rows').length === 0)
  })

  it("asks Value Field Settings with the value field's index in the PivotTable, not in a deferred layout", async () => {
    const { w } = await report()
    const defer = document.querySelector<HTMLInputElement>('[data-xs-pf-defer]')!
    defer.checked = true
    defer.dispatchEvent(new Event('change'))
    await setBox(3, true)
    expect(chips('values')).toEqual(['Sum of Amount', 'Year'])
    await chipMenu('values', 1)
    expect(item('valueFieldSettings')?.disabled).toBe(true)
    item('moveUp')!.click()
    await settle()
    expect(chips('values')).toEqual(['Year', 'Sum of Amount'])
    await chipMenu('values', 1)
    item('valueFieldSettings')!.click()
    await settle()
    // the saved PivotTable's first value field, Sum of Amount, whatever the deferred layout's order
    await until(() => document.querySelector<HTMLInputElement>('[data-pv-name]')?.value === 'Sum of Amount')
    expect(w.emitted('pivotDialog')).toBeUndefined()
  })

  it('a drag from the list ends without toggling the box, on Esc anywhere, and only for its own pointer', async () => {
    await report()
    const row = document.querySelector<HTMLElement>('[data-xs-pf-field="2"]')!
    const at = (x: number, id = 1) => ({ clientX: x, clientY: x, pointerId: id, bubbles: true, button: 0 })
    const ghost = () => document.querySelector('.xs-pf-ghost')
    row.dispatchEvent(new PointerEvent('pointerdown', at(5)))
    window.dispatchEvent(new PointerEvent('pointermove', at(60)))
    await settle()
    expect(ghost()).not.toBeNull()
    window.dispatchEvent(new PointerEvent('pointerup', at(60, 2)))
    await settle()
    expect(ghost()).not.toBeNull()
    document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    await settle()
    expect(ghost()).toBeNull()
    row.dispatchEvent(new PointerEvent('pointerdown', at(5)))
    window.dispatchEvent(new PointerEvent('pointermove', at(60)))
    window.dispatchEvent(new PointerEvent('pointerup', at(5)))
    row.querySelector('label')!.click()
    await settle()
    expect(box(2).checked).toBe(false)
    expect(chips('rows')).toEqual(['Region'])
  })
})

describe('Insert PivotTable', () => {
  async function data(book = pivotBook()): Promise<VueWrapper> {
    const s = localSource(book, { calc: 'inline' })
    sources.push(s)
    const w = await mountSheet(s, { style: 'height: 400px' })
    await goTo(w, 'Data!B2')
    return w
  }
  const dialog = () => document.querySelector<HTMLElement>('[data-xs-pivot-dialog]')
  const field = (name: string) => document.querySelector<HTMLInputElement>(`[data-xs-pivot-${name}]`)!

  it('proposes the current region with its sheet and puts the PivotTable on a new sheet at A3, with the pane', async () => {
    const w = await data()
    expect(lastState(w)?.canInsertPivot).toBe(true)
    expect(await api(w).insertPivotTable()).toBe(true)
    await settle()
    expect(dialog()?.querySelector('.xs-modal-title')?.textContent).toBe('PivotTable from table or range')
    expect(field('source').value).toBe('Data!$A$1:$D$4')
    document.querySelector<HTMLButtonElement>('[data-xs-dialog-ok]')!.click()
    await until(() => lastState(w)?.sheet === 'Sheet3' && lastState(w)?.ref === 'A3')
    expect(dialog()).toBeNull()
    // xlcalc names it as the first free PivotTable number in the workbook
    expect(lastState(w)?.pivot).toBe('PivotTable2')
    await until(() => !!pane())
    expect([0, 1, 2, 3].map((i) => box(i).checked)).toEqual([false, false, false, false])
    expect([...document.querySelectorAll('[role="tab"]')].map((x) => x.textContent?.trim())).toEqual(['Sheet3', 'Data', 'Report'])
  })

  it('proposes nothing for an empty cell alone, whatever data lies elsewhere', async () => {
    const w = await data(pivotBook({ extra: { H12: { value: 1 } } }))
    // within the sheet's data, with the PivotTable and H12 read around it
    await goTo(w, 'Report!E9')
    await api(w).insertPivotTable()
    await settle()
    expect(field('source').value).toBe('')
  })

  it("puts it at an existing sheet's location, and keeps the dialog with Excel's alert for a source or a place it does not take", async () => {
    const w = await data()
    await api(w).insertPivotTable()
    await settle()
    field('source').value = 'A1:'
    field('source').dispatchEvent(new Event('input'))
    document.querySelector<HTMLButtonElement>('[data-xs-dialog-ok]')!.click()
    await settle()
    expect(document.querySelector('[data-xs-pivot-error]')?.textContent).toBe('Data source reference is not valid.')
    field('source').value = 'Data!A1:D4'
    field('source').dispatchEvent(new Event('input'))
    const existing = document.querySelector<HTMLInputElement>('[data-xs-pivot-place="existing"]')!
    existing.checked = true
    existing.dispatchEvent(new Event('change'))
    await settle()
    field('location').value = 'Nowhere!F3'
    field('location').dispatchEvent(new Event('input'))
    document.querySelector<HTMLButtonElement>('[data-xs-dialog-ok]')!.click()
    await settle()
    expect(document.querySelector('[data-xs-pivot-error]')?.textContent).toBe('Destination reference is not valid.')
    field('location').value = 'Report!$F$3'
    field('location').dispatchEvent(new Event('input'))
    document.querySelector<HTMLButtonElement>('[data-xs-dialog-ok]')!.click()
    await until(() => lastState(w)?.sheet === 'Report' && lastState(w)?.ref === 'F3')
    expect(lastState(w)?.pivot).toBe('PivotTable2')
    expect(lastState(w)?.fieldList).toBe(true)
  })
})
