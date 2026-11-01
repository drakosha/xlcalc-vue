// Slicers and timelines on SheetView over localSource: drawn at their places with their items and periods; a click, Ctrl+click,
// Multi-Select and Clear Filter filter the PivotTable; a run of periods; selecting, Delete and the menu; Insert Slicers,
// Insert Timelines, Slicer Settings and Report Connections; the protected sheet; frozen panes; the timeline's labels.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { enMessages, ruMessages, type SheetViewApi, type SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { selectionLabel, timelinePeriods } from '../src/timeline'
import { cellText, goTo, mountSheet, settle, until } from './helpers'
import { pivotBook } from './pivot-fixture'
import { slicerBook } from './slicer-fixture'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

async function report(book = slicerBook()): Promise<{ w: VueWrapper; s: LocalSource; api: SheetViewApi }> {
  const s = localSource(book, { calc: 'inline' })
  sources.push(s)
  const w = await mountSheet(s)
  await goTo(w, 'Report!B4')
  await until(() => !!document.querySelector('[data-xs-object="slicer:Region"]'))
  return { w, s, api: w.vm as unknown as SheetViewApi }
}
const q = <T extends Element = HTMLElement>(sel: string) => document.querySelector<T>(sel)
const item = (n: number) => q<HTMLButtonElement>(`[data-xs-slicer-item="${n}"]`)!
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
const rows = (w: VueWrapper) => [4, 5, 6].map((r) => `${cellText(w, r, 1)}=${cellText(w, r, 2)}`)
async function click(el: Element, init: MouseEventInit = {}): Promise<void> {
  el.dispatchEvent(new MouseEvent('click', { bubbles: true, ...init }))
  await settle()
}
async function press(el: Element, type: string, init: MouseEventInit = {}): Promise<void> {
  el.dispatchEvent(new MouseEvent(type, { bubbles: true, button: 0, ...init }))
  await settle()
}

describe('slicers and timelines on the sheet', () => {
  it('draws the slicer with its items and the timeline with its periods at their places', async () => {
    await report()
    const slicer = q('[data-xs-object="slicer:Region"]')!
    expect(slicer.querySelector('.xs-slicer-caption')!.textContent).toBe('Region')
    expect([...slicer.querySelectorAll('[data-xs-slicer-item]')].map((x) => [x.textContent!.trim(), x.getAttribute('aria-selected')])).toEqual([
      ['North', 'true'],
      ['South', 'true'],
    ])
    // D3: three columns of 71 px after the 56 px numbers, two rows of 20 px under the 24 px letters
    expect(slicer.style.left).toBe(`${56 + 3 * 71}px`)
    expect(slicer.style.top).toBe(`${24 + 2 * 20}px`)
    const timeline = q('[data-xs-object="timeline:Date"]')!
    expect(timeline.querySelector('[data-xs-timeline-label]')!.textContent).toBe('All Periods')
    expect(timeline.querySelector('[data-xs-timeline-level]')!.textContent!.trim()).toBe('QUARTERS')
    expect([...timeline.querySelectorAll('[data-xs-period]')].map((x) => x.textContent!.trim())).toEqual(['Q1', 'Q2', 'Q3', 'Q4', 'Q1', 'Q2', 'Q3', 'Q4'])
  })

  it('a click chooses an item alone, Ctrl+click adds one, Clear Filter chooses all again', async () => {
    const { w } = await report()
    await click(item(1))
    await until(() => cellText(w, 4, 1) === 'South')
    expect(item(0).getAttribute('aria-selected')).toBe('false')
    await click(item(0), { ctrlKey: true })
    await until(() => cellText(w, 5, 1) === 'South')
    expect(rows(w)).toEqual(['North=15', 'South=20', 'Grand Total=35'])
    await click(item(0), { ctrlKey: true })
    await until(() => cellText(w, 4, 1) === 'South')
    await click(q('[data-xs-slicer-clear]')!)
    await until(() => cellText(w, 4, 1) === 'North')
    expect(q<HTMLButtonElement>('[data-xs-slicer-clear]')!.disabled).toBe(true)
  })

  it('with Multi-Select a click takes an item away; taking away the last one clears the filter', async () => {
    const { w } = await report()
    await click(q('[data-xs-slicer-multi]')!)
    expect(q('[data-xs-slicer-multi]')!.getAttribute('aria-pressed')).toBe('true')
    await click(item(0))
    await until(() => cellText(w, 4, 1) === 'South')
    await click(item(1))
    await until(() => cellText(w, 4, 1) === 'North')
    expect(rows(w)).toEqual(['North=15', 'South=20', 'Grand Total=35'])
  })

  it('a press on a period of the timeline chooses it: the PivotTable, the label and the filled period follow', async () => {
    const { w } = await report()
    const grid = q('[data-xs-periods]')!
    await press(q('[data-xs-period="0"]')!, 'pointerdown')
    await press(grid, 'pointerup')
    await until(() => cellText(w, 5, 1) === 'Grand Total')
    expect(rows(w)).toEqual(['North=10', 'Grand Total=10', '='])
    expect(q('[data-xs-timeline-label]')!.textContent).toBe('Q1 2024')
    expect(q('[data-xs-period="0"]')!.getAttribute('aria-selected')).toBe('true')
    await click(q('[data-xs-timeline-clear]')!)
    await until(() => q('[data-xs-timeline-label]')!.textContent === 'All Periods')
    await click(q('[data-xs-timeline-level]')!)
    await click(q('[data-xs-level="years"]')!)
    await until(() => document.querySelectorAll('[data-xs-period]').length === 2)
  })

  it('a press selects a slicer: the state names it, Delete removes it', async () => {
    const { w } = await report()
    const box = q('[data-xs-object="slicer:Region"]')!
    await press(box.querySelector('.xs-slicer-head')!, 'pointerdown')
    expect(lastState(w)?.object).toEqual({ kind: 'slicer', name: 'Region' })
    expect(box.querySelectorAll('[data-xs-handle]').length).toBe(8)
    box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }))
    await until(() => !q('[data-xs-object="slicer:Region"]'))
    expect(lastState(w)?.object).toBeUndefined()
  })

  it("the slicer's menu has Excel's items; Remove removes it", async () => {
    await report()
    const box = q('[data-xs-object="slicer:Region"]')!
    box.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 300, clientY: 100 }))
    await settle()
    expect([...document.querySelectorAll('[data-xs-menu]')].map((x) => x.textContent!.trim())).toEqual(['Refresh', 'Sort A to Z', 'Sort Z to A', 'Clear Filter from "Region"', 'Multi-Select "Region"', 'Report Connections...', 'Remove "Region"', 'Slicer Settings...'])
    expect(q<HTMLButtonElement>('[data-xs-menu="objectClear"]')!.disabled).toBe(true)
    await click(q('[data-xs-menu="objectRemove"]')!)
    await until(() => !q('[data-xs-object="slicer:Region"]'))
  })

  it('a drag of the frame moves the slicer to a new place, as one operation', async () => {
    const { s } = await report()
    const head = q('[data-xs-object="slicer:Region"] .xs-slicer-head')!
    await press(head, 'pointerdown', { clientX: 300, clientY: 100 })
    await press(head, 'pointermove', { clientX: 371, clientY: 140 })
    await press(head, 'pointerup', { clientX: 371, clientY: 140 })
    await until(() => q<HTMLElement>('[data-xs-object="slicer:Region"]')?.style.left === `${56 + 4 * 71}px`)
    expect((await s.sheets())[1]!.slicers![0]!.place).toMatchObject({ from: 'E5', to: 'G14' })
  })

  it('Insert Slicers adds a slicer of each field checked and selects it; Insert Timelines offers the fields of dates', async () => {
    const { w, api } = await report()
    expect(lastState(w)?.canInsertSlicer).toBe(true)
    expect(api.insertSlicer()).toBe(true)
    await settle()
    expect(q('.xs-modal-title')!.textContent).toBe('Insert Slicers')
    expect([...document.querySelectorAll('[data-xs-field]')].map((x) => x.getAttribute('data-xs-field'))).toEqual(['Region', 'Amount', 'Kind', 'Year', 'Date'])
    // OK is unavailable until a field is checked, as in Excel
    expect(q<HTMLButtonElement>('[data-xs-dialog-ok]')!.disabled).toBe(true)
    await click(q('[data-xs-field="Kind"]')!)
    await click(q('[data-xs-dialog-ok]')!)
    await until(() => !!q('[data-xs-object="slicer:Kind"]'))
    await until(() => lastState(w)?.object?.name === 'Kind')
    expect(api.insertTimeline()).toBe(true)
    await settle()
    expect(q('.xs-modal-title')!.textContent).toBe('Insert Timelines')
    expect([...document.querySelectorAll('[data-xs-field]')].map((x) => x.getAttribute('data-xs-field'))).toEqual(['Date'])
    await goTo(w, 'Q20')
    expect(api.insertSlicer()).toBe(false)
  })

  it("Insert Timelines over a PivotTable without dates shows Excel's alert", async () => {
    const s = localSource(pivotBook(), { calc: 'inline' })
    sources.push(s)
    const w = await mountSheet(s)
    await goTo(w, 'Report!B4')
    await until(() => lastState(w)?.canInsertSlicer === true)
    expect((w.vm as unknown as SheetViewApi).insertTimeline()).toBe(true)
    await until(() => !!q('[data-xs-alert-message]'))
    expect(q('[data-xs-alert-message]')!.textContent).toBe("We can't create a Timeline for this report because it doesn't have a field formatted as Date.")
  })

  it("Slicer Settings changes the caption and the sorting; Report Connections lists the cache's PivotTables", async () => {
    const { api } = await report()
    await press(q('[data-xs-object="slicer:Region"] .xs-slicer-head')!, 'pointerdown')
    expect(api.slicerSettings()).toBe(true)
    await settle()
    expect(q('.xs-modal-title')!.textContent).toBe('Slicer Settings')
    const caption = q<HTMLInputElement>('[data-xs-slicer-caption]')!
    caption.value = 'Regions'
    caption.dispatchEvent(new Event('input'))
    await click(q('[data-xs-slicer-descending]')!)
    await click(q('[data-xs-dialog-ok]')!)
    await until(() => q('[data-xs-object="slicer:Region"] .xs-slicer-caption')?.textContent === 'Regions')
    expect([...document.querySelectorAll('[data-xs-slicer-item]')].map((x) => x.textContent!.trim())).toEqual(['South', 'North'])
    expect(api.reportConnections()).toBe(true)
    await settle()
    expect(q('.xs-modal-title')!.textContent).toBe('Report Connections (Region)')
    expect(q<HTMLInputElement>('[data-xs-connection="Report!PivotTable1"]')!.checked).toBe(true)
  })

  it("Slicer Settings renames the slicer, keeping it selected; a name a timeline has is xlcalc's refusal", async () => {
    const { w, api } = await report()
    const rename = async (to: string) => {
      await press(q(`[data-xs-object^="slicer:"] .xs-slicer-head`)!, 'pointerdown')
      expect(api.slicerSettings()).toBe(true)
      await settle()
      const name = q<HTMLInputElement>('[data-xs-slicer-name]')!
      expect(name.disabled).toBe(false)
      name.value = to
      name.dispatchEvent(new Event('input'))
      await click(q('[data-xs-dialog-ok]')!)
    }
    await rename('Regions')
    await until(() => !!q('[data-xs-object="slicer:Regions"]'))
    expect(lastState(w)?.object).toEqual({ kind: 'slicer', name: 'Regions' })
    await rename('Date')
    await until(() => !!q('[data-xs-alert-message]'))
    expect(q('[data-xs-object="slicer:Regions"]')).not.toBeNull()
  })

  it('Slicer Settings never sends an empty name: OK waits for one and the dialog stays with the other settings', async () => {
    const { s, api } = await report()
    await press(q('[data-xs-object="slicer:Region"] .xs-slicer-head')!, 'pointerdown')
    expect(api.slicerSettings()).toBe(true)
    await settle()
    const caption = q<HTMLInputElement>('[data-xs-slicer-caption]')!
    caption.value = 'Regions'
    caption.dispatchEvent(new Event('input'))
    const name = q<HTMLInputElement>('[data-xs-slicer-name]')!
    name.value = ' '
    name.dispatchEvent(new Event('input'))
    await settle()
    expect(q<HTMLButtonElement>('[data-xs-dialog-ok]')!.disabled).toBe(true)
    name.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    await settle()
    expect(q('[data-xs-slicer-dialog]')).not.toBeNull()
    expect(q('[data-xs-alert-message]')).toBeNull()
    expect((await s.sheets())[1]!.slicers![0]).toMatchObject({ name: 'Region', caption: 'Region' })
    name.value = 'Regions'
    name.dispatchEvent(new Event('input'))
    await settle()
    await click(q('[data-xs-dialog-ok]')!)
    await until(() => q('[data-xs-object="slicer:Regions"] .xs-slicer-caption')?.textContent === 'Regions')
  })

  it('on a protected sheet the items are disabled without Use PivotTable and PivotChart, and the handles without Edit objects', async () => {
    await report(slicerBook({ protection: {} }))
    expect(item(0).disabled).toBe(true)
    document.body.innerHTML = ''
    await report(slicerBook({ protection: { pivotTables: true } }))
    expect(item(0).disabled).toBe(false)
    await press(q('[data-xs-object="slicer:Region"] .xs-slicer-head')!, 'pointerdown')
    expect(document.querySelectorAll('[data-xs-handle]').length).toBe(0)
  })

  it('with frozen rows the panes draw their parts of a box across them', async () => {
    await report(slicerBook({ freeze: true }))
    expect([...document.querySelectorAll('[data-xs-object="slicer:Region"]')].map((x) => x.getAttribute('data-xs-part') ?? '')).toEqual(['r', ''])
  })
})

describe("a timeline's periods and label", () => {
  const bounds = { start: '2024-01-01T00:00:00', end: '2025-01-01T00:00:00' }
  it('lists the periods of a level between the bounds, under their years or months', () => {
    expect(timelinePeriods(bounds, 'months', enMessages).map((p) => p.label)).toEqual(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
    expect(timelinePeriods(bounds, 'quarters', ruMessages)[1]).toEqual({ start: '2024-04-01T00:00:00', end: '2024-06-30T00:00:00', label: 'К2', group: '2024' })
    expect(timelinePeriods(bounds, 'days', enMessages)).toHaveLength(366)
    expect(timelinePeriods(bounds, 'days', enMessages)[31]).toMatchObject({ label: '1', group: 'Feb 2024' })
  })
  it("words the selection in Excel's patterns", () => {
    const s = (start: string, end: string) => ({ start: `${start}T00:00:00`, end: `${end}T00:00:00` })
    expect(selectionLabel(undefined, 'months', enMessages)).toBe('All Periods')
    expect(selectionLabel(s('2024-03-01', '2024-05-31'), 'months', enMessages)).toBe('Mar - May 2024')
    expect(selectionLabel(s('2023-11-01', '2024-02-29'), 'months', enMessages)).toBe('Nov 2023 - Feb 2024')
    expect(selectionLabel(s('2024-01-01', '2025-12-31'), 'years', ruMessages)).toBe('2024–2025 гг.')
    expect(selectionLabel(s('2024-03-15', '2024-03-15'), 'days', enMessages)).toBe('Mar 15, 2024')
    expect(selectionLabel(s('2024-04-01', '2024-06-30'), 'quarters', ruMessages)).toBe('К2 2024 г.')
  })
})

describe('the texts of slicers and timelines', () => {
  it("are Excel for Windows' as measured, in English and Russian", () => {
    const words = (m: typeof enMessages) => [m.menuMultiSelect('Region'), m.menuClearFilterFrom('Region'), m.menuRemoveObject('Region'), m.menuSlicerSettings, m.menuReportConnections, m.slicerClearFilter, m.insertSlicersTitle, m.insertTimelinesTitle, m.timelineLevel('quarters')]
    expect(words(enMessages)).toEqual(['Multi-Select "Region"', 'Clear Filter from "Region"', 'Remove "Region"', 'Slicer Settings...', 'Report Connections...', 'Clear Filter (Alt+C)', 'Insert Slicers', 'Insert Timelines', 'QUARTERS'])
    expect(words(ruMessages)).toEqual(['Выбор нескольких объектов "Region"', 'Снять фильтрацию с: "Region"', 'Удалить "Region"', 'Настройка среза...', 'Подключения к отчетам...', 'Очистить фильтр (Alt+C)', 'Вставка срезов', 'Вставка временных шкал', 'КВАРТАЛЫ'])
  })
})
