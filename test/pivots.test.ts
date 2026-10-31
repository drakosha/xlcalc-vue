// PivotTables on SheetView over localSource: the state names the PivotTable under the active cell, Alt+F5, the cell menu's
// Refresh and refreshPivot() refresh it through the source's structure, and the source's refusal shows as Excel's alert.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import { SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { cellText, editorKey, editorOpen, goTo, key, mountSheet, settle, type, until } from './helpers'
import { addSales, pivotBook, salesBook } from './pivot-fixture'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

function open(book = pivotBook()): LocalSource {
  const s = localSource(book, { calc: 'inline' })
  sources.push(s)
  return s
}
/** The sheet over Report (the second sheet), with the source's Data!B2 changed so that a refresh shows. */
async function report(book = pivotBook()): Promise<{ w: VueWrapper; s: LocalSource }> {
  const s = open(book)
  await s.edit([{ sheet: 'Data', cell: 'B2', text: '100' }])
  const w = await mountSheet(s)
  await goTo(w, 'Report!B4')
  await until(() => cellText(w, 4, 2) === '15')
  return { w, s }
}
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
const menuItem = (id: string) => document.querySelector<HTMLButtonElement>(`[data-xs-menu="${id}"]`)
async function context(w: VueWrapper, cell: string): Promise<void> {
  await w.find(`[data-cell="${cell}"]`).trigger('contextmenu', { button: 2, clientX: 40, clientY: 60 })
  await settle()
}
const alertText = () => document.querySelector('[data-xs-alert="notice"] [data-xs-alert-message]')?.textContent ?? null

describe('PivotTables on the sheet', () => {
  it('the state names the PivotTable under the active cell, and Alt+F5 refreshes it', async () => {
    const { w } = await report()
    expect(lastState(w)).toMatchObject({ pivot: 'PivotTable1', canRefreshPivot: true })
    await key(w, 'F5', { altKey: true })
    await until(() => cellText(w, 4, 2) === '105')
    await goTo(w, 'D10')
    expect(lastState(w)?.pivot).toBeUndefined()
    expect(lastState(w)?.canRefreshPivot).toBeUndefined()
    expect(await (w.vm as unknown as SheetViewApi).refreshPivot()).toBeNull()
  })

  it("the cell menu of a PivotTable's cell has Refresh, and refreshPivot() refreshes it", async () => {
    const { w } = await report()
    await context(w, '5:1')
    expect(menuItem('refreshPivot')?.textContent?.trim()).toBe('Refresh')
    menuItem('refreshPivot')!.click()
    await until(() => cellText(w, 4, 2) === '105')
    await context(w, '9:4')
    expect(menuItem('refreshPivot')).toBeNull()
    await goTo(w, 'B6')
    expect(await (w.vm as unknown as SheetViewApi).refreshPivot()).toMatchObject({ ok: true })
  })

  it('refuses Refresh on a protected sheet with or without Use PivotTable and PivotChart, as Excel for Windows does', async () => {
    for (const [protection, said] of [
      [{}, 'Cannot edit PivotTable on protected sheet.'],
      [{ pivotTables: true }, /^That command cannot be performed while a protected sheet contains another PivotTable report/],
    ] as const) {
      const { w } = await report(pivotBook({ protection }))
      expect(lastState(w)?.pivot).toBe('PivotTable1')
      expect(lastState(w)?.canRefreshPivot).toBeUndefined()
      await context(w, '4:2')
      expect(menuItem('refreshPivot')?.disabled).toBe(true)
      document.body.click()
      // Alt+F5 does nothing, without an alert
      await key(w, 'F5', { altKey: true })
      await settle()
      expect(alertText()).toBeNull()
      expect(await (w.vm as unknown as SheetViewApi).refreshPivot()).toBeNull()
      // Ctrl+Alt+F5 (Refresh All) shows the source's refusal as Excel's alert
      await key(w, 'F5', { altKey: true, ctrlKey: true })
      await until(() => alertText() !== null)
      expect(alertText()).toMatch(said)
      document.querySelector<HTMLButtonElement>('[data-xs-alert-button="ok"]')!.click()
      await settle()
      expect(cellText(w, 4, 2)).toBe('15')
      w.unmount()
    }
  })

  it("typing or F2 on a value shows Excel's alert at once and opens no editor; over a label typing renames", async () => {
    const { w } = await report()
    for (const k of ['9', 'F2']) {
      await goTo(w, 'B4')
      await key(w, k)
      await settle()
      expect(alertText()).toBe("We can't change this part of the PivotTable.")
      expect(editorOpen(w)).toBe(false)
      document.querySelector<HTMLButtonElement>('[data-xs-alert-button="ok"]')!.click()
      await settle()
      expect(alertText()).toBeNull()
    }
    expect(cellText(w, 4, 2)).toBe('15')
    await goTo(w, 'A5')
    await key(w, 'F2')
    await type(w, 'x')
    await editorKey(w, 'Enter')
    await until(() => cellText(w, 5, 1) === 'x')
  })

  it('without value fields: the labels in its values area take typing, its empty cells refuse at the first key', async () => {
    const s = open(salesBook())
    expect(await s.structure(addSales({ rows: [0], columns: [1] }))).toMatchObject({ ok: true })
    const pivots = (await s.sheets())[1]!.pivots!
    // xlcalc's areas: the values area is Excel's DataBodyRange, empty cells right of the labels
    expect(pivots[0]).toMatchObject({ range: 'A3:D8', headers: 'A3:D4', rowLabels: 'A5:A8', values: 'B5:D8' })
    const w = await mountSheet(s)
    await goTo(w, 'R!C6')
    await key(w, '9')
    await settle()
    expect(alertText()).toBe("We can't change this part of the PivotTable.")
    document.querySelector<HTMLButtonElement>('[data-xs-alert-button="ok"]')!.click()
    await settle()
    await goTo(w, 'A5')
    await key(w, 'F2')
    expect(alertText()).toBeNull()
    expect(editorOpen(w)).toBe(true)
  })

  it('a PivotTable with row fields alone has its labels as its values area, and typing over them renames', async () => {
    const s = open(salesBook())
    expect(await s.structure(addSales({ rows: [0] }))).toMatchObject({ ok: true })
    expect((await s.sheets())[1]!.pivots![0]).toMatchObject({ rowLabels: 'A4:A7', values: 'A4:A7' })
    const w = await mountSheet(s)
    await goTo(w, 'R!A4')
    await key(w, 'F2')
    expect(alertText()).toBeNull()
    expect(editorOpen(w)).toBe(true)
  })

  it("the toolbar's Refresh All list has Refresh, enabled on a PivotTable's cell", async () => {
    const s = open()
    await s.edit([{ sheet: 'Data', cell: 'B2', text: '100' }])
    const sheet = ref<SheetViewApi | null>(null)
    const state = ref<SheetViewState | null>(null)
    const Host = defineComponent(() => () => [h(SheetToolbar, { source: s, sheet: sheet.value, state: state.value }), h(SheetView, { ref: sheet, source: s, style: 'height: 400px', onState: (x: SheetViewState | null) => (state.value = x) })])
    const w = mount(Host, { attachTo: document.body })
    await until(() => !!sheet.value && !!state.value)
    const view = w.findComponent(SheetView)
    await goTo(view, 'Report!B4')
    await until(() => state.value?.canRefreshPivot === true)
    await w.find('[data-tb="refresh-all-menu"]').trigger('click')
    const item = w.find('[data-tb="refresh-pivot"]')
    expect(item.text()).toBe('Refresh')
    expect(item.attributes('title')).toBe('Refresh (Alt+F5)')
    await item.trigger('click')
    await until(() => cellText(view, 4, 2) === '105')
    await goTo(view, 'D10')
    await until(() => state.value?.canRefreshPivot === undefined)
    await w.find('[data-tb="refresh-all-menu"]').trigger('click')
    expect(w.find('[data-tb="refresh-pivot"]').attributes('disabled')).toBeDefined()
    w.unmount()
  })
})
