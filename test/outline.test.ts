// Excel's outline on the grid: the bars beside the numbers and above the letters with their brackets and buttons, the
// buttons and level buttons through the source's `outline`, Group and Ungroup from the menu, the keys and the box that
// asks rows or columns, and the sheet's room for the bars.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, ruMessages, type MemoryBook, type SheetViewApi } from '../src'
import type { EditOutcome, OutlineChange } from '../src/contract/source'
import { address, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
  vi.restoreAllMocks()
})

/** Rows 2:4 and 6:9 (collapsed, under row 10) inside 2:10; columns B:D inside B:E, the summary column F; row 1 frozen. */
const book = (): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: Object.fromEntries(Array.from({ length: 12 }, (_, i) => [`A${i + 1}`, i === 0 ? 'Name' : i % 2 ? 'a' : 'b'])),
      heights: { '6:9': 0 },
      freeze: { rows: 1, cols: 0 },
      outline: { rows: { '2:4': 2, 5: 1, '6:9': 2, 10: 1 }, columns: { 'B:D': 2, E: 1 }, collapsed: ['10'] },
    },
    { name: 'T', cells: { A1: 1 } },
  ],
})
const OK: EditOutcome = { ok: true, changes: [] }
/** The memory book with its `outline` watched. */
function watched() {
  const source = memorySource(book())
  const outline = vi.spyOn(source, 'outline')
  return { source, outline }
}
const numbers = (w: VueWrapper) => w.findAll('.xs-num').map((n) => n.text())
const rowBar = (w: VueWrapper, row: number) => w.findAll('tbody tr').find((tr) => tr.find('.xs-num').text() === String(row))!.find('.xs-ol-rows')
const button = (w: VueWrapper, axis: 'rows' | 'columns', id: string) => w.find(`.xs-ol-${axis} [data-xs-outline="${id}"]`)
const level = (w: VueWrapper, axis: 'rows' | 'columns', k: number) => w.find(`.xs-ol-levels-${axis} [data-xs-outline-level="${k}"]`)
const sent = (spy: { mock: { calls: unknown[][] } }) => spy.mock.calls.map((c) => c[0] as OutlineChange)
const items = () => [...document.querySelectorAll<HTMLElement>('[data-xs-menu]')].map((b) => b.dataset.xsMenu)
async function menuOn(el: { trigger: (e: string, o?: object) => Promise<void> }, pick: string): Promise<void> {
  await el.trigger('contextmenu', { button: 2, clientX: 40, clientY: 60 })
  await settle()
  document.querySelector<HTMLElement>(`[data-xs-menu="${pick}"]`)!.click()
  await settle()
}

describe('the outline bars', () => {
  it('draw brackets with a tick at the far end and a button on the summary line, − expanded and + collapsed', async () => {
    const w = await mountSheet(memorySource(book()))
    expect(numbers(w).slice(0, 7)).toEqual(['1', '2', '3', '4', '5', '10', '11'])
    const slots = (row: number) =>
      rowBar(w, row)
        .findAll('.xs-ol-mark')
        .map((mk) => [mk.find('.xs-ol-tick').exists() ? 'tick' : '', mk.find('.xs-ol-line').exists() ? 'line' : '', mk.find('.xs-ol-end').exists() ? 'end' : '', (mk.find('button').exists() ? mk.find('button').classes().find((c) => c === 'xs-ol-minus' || c === 'xs-ol-plus') : '')].filter(Boolean).join(' '))
    expect([1, 2, 3, 5, 10, 11].map(slots)).toEqual([[], ['tick line', 'tick line'], ['line', 'line'], ['line', 'end xs-ol-minus'], ['line', 'xs-ol-plus'], ['end xs-ol-minus']])
    expect(button(w, 'rows', '6:9:2').attributes('title')).toBe('Show Detail')
    expect(button(w, 'rows', '2:4:2').attributes('aria-expanded')).toBe('true')
    // Columns: B:E at level 1 with its button on F, B:D at level 2 with its button on E.
    const cols = w.findAll('.xs-ol-columns')
    expect(cols[1]!.findAll('.xs-ol-tick')).toHaveLength(2)
    expect(button(w, 'columns', '2:4:2').exists()).toBe(true)
    expect(button(w, 'columns', '2:5:1').exists()).toBe(true)
    expect(w.findAll('.xs-ol-levels-rows button').map((b) => b.text())).toEqual(['1', '2', '3'])
    expect(w.findAll('.xs-ol-levels-columns button').map((b) => b.text())).toEqual(['1', '2', '3'])
  })

  it('take three slots of 16 px each way, the headers and numbers stick past them, frozen rows below both', async () => {
    const w = await mountSheet(memorySource(book()))
    const table = w.find('.xs-table').element as HTMLElement
    expect([table.style.getPropertyValue('--xs-outline-w'), table.style.getPropertyValue('--xs-outline-h')]).toEqual(['48px', '48px'])
    expect((w.find('[data-cell="1:1"]').element as HTMLElement).style.top).toBe(`${24 + 48}px`)
    expect(w.find('.xs-ol-head').attributes('style')).toContain('height: 48px')
    const plain = await mountSheet(memorySource({ sheets: [{ name: 'P', cells: { A1: 1 } }] }))
    expect(plain.find('.xs-ol').exists()).toBe(false)
    expect(plain.find('.xs-ol-levels').exists()).toBe(false)
    expect((plain.find('.xs-table').element as HTMLElement).style.getPropertyValue('--xs-outline-w')).toBe('0px')
  })

  it('leave out the rows’ bar under a display filter, the columns’ stays', async () => {
    localStorage.setItem('k:filter:S', JSON.stringify({ cols: { 1: { values: ['a'] } }, sort: null }))
    const w = await mountSheet(memorySource(book()), { storageKey: 'k' })
    expect(w.find('[data-xs-filter-status]').exists()).toBe(true)
    expect(w.find('.xs-ol-rows').exists()).toBe(false)
    expect(w.find('.xs-ol-levels-rows').exists()).toBe(false)
    expect(w.find('.xs-ol-columns').exists()).toBe(true)
  })
})

describe('the buttons', () => {
  it('a group’s button collapses or expands it through the source, and the sheet shows the result', async () => {
    const { source, outline } = watched()
    const w = await mountSheet(source)
    await button(w, 'rows', '6:9:2').trigger('click')
    await settle()
    expect(numbers(w).slice(0, 10)).toEqual(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'])
    await button(w, 'rows', '2:10:1').trigger('click')
    await settle()
    expect(numbers(w).slice(0, 3)).toEqual(['1', '11', '12'])
    await button(w, 'columns', '2:4:2').trigger('click')
    await settle()
    expect(sent(outline)).toEqual([
      { op: 'expand', sheet: 'S', level: 2, rows: '6:9' },
      { op: 'collapse', sheet: 'S', level: 1, rows: '2:10' },
      { op: 'collapse', sheet: 'S', level: 2, columns: 'B:D' },
    ])
    expect(w.findAll('.xs-head').map((h) => h.text()).slice(0, 3)).toEqual(['A', 'E', 'F'])
  })

  it('a level button shows that level: 1 the summaries alone, 3 everything', async () => {
    const { source, outline } = watched()
    const w = await mountSheet(source)
    await level(w, 'rows', 1).trigger('click')
    await settle()
    expect(numbers(w).slice(0, 3)).toEqual(['1', '11', '12'])
    await level(w, 'rows', 3).trigger('click')
    await settle()
    expect(numbers(w).slice(0, 11)).toEqual(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11'])
    await level(w, 'columns', 2).trigger('click')
    expect(sent(outline)).toEqual([
      { op: 'showLevel', sheet: 'S', axis: 'rows', level: 1 },
      { op: 'showLevel', sheet: 'S', axis: 'rows', level: 3 },
      { op: 'showLevel', sheet: 'S', axis: 'columns', level: 2 },
    ])
  })

  it('read the sheet again after a source that does not tell; a refusal shows its message', async () => {
    const source = memorySource(book())
    const sheets = vi.spyOn(source, 'sheets')
    source.outline = vi.fn(async (): Promise<EditOutcome> => ({ ok: false, reason: 'refused', message: 'Нельзя' }))
    const w = await mountSheet(source)
    const before = sheets.mock.calls.length
    await button(w, 'rows', '2:4:2').trigger('click')
    await settle()
    expect(w.find('.xs-message').text()).toBe('Нельзя')
    expect(sheets.mock.calls.length).toBe(before)
    source.outline = vi.fn(async () => OK)
    await button(w, 'rows', '2:4:2').trigger('click')
    await settle()
    expect(sheets.mock.calls.length).toBe(before + 1)
  })

  it('are drawn but do not press without the source’s outline, or on a read-only sheet', async () => {
    const plain = memorySource(book())
    delete (plain as { outline?: unknown }).outline
    const w = await mountSheet(plain)
    expect(button(w, 'rows', '6:9:2').attributes('disabled')).toBeDefined()
    expect(level(w, 'rows', 1).attributes('disabled')).toBeDefined()
    const { source, outline } = watched()
    const ro = await mountSheet(source, { readOnly: true })
    expect(button(ro, 'rows', '6:9:2').attributes('disabled')).toBeDefined()
    await button(ro, 'rows', '6:9:2').trigger('click')
    expect(outline).not.toHaveBeenCalled()
  })

  it('SheetView exposes outline, as the buttons go; null without the source’s method', async () => {
    const { source, outline } = watched()
    const w = await mountSheet(source)
    expect(await (w.vm as unknown as SheetViewApi).outline({ op: 'expand', sheet: 'S', rows: '6:9', level: 2 })).toEqual(OK)
    await settle()
    expect(numbers(w)).toContain('7')
    expect(outline).toHaveBeenCalledTimes(1)
    delete (source as { outline?: unknown }).outline
    const bare = await mountSheet(source)
    expect(await (bare.vm as unknown as SheetViewApi).outline({ op: 'showLevel', sheet: 'S', axis: 'rows', level: 1 })).toBeNull()
  })
})

describe('Group and Ungroup', () => {
  const num = (w: VueWrapper, row: number) => w.findAll('.xs-num').find((n) => n.text() === String(row))!
  const head = (w: VueWrapper, col: number) => w.findAll('.xs-head')[col - 1]!

  it('from the menu on numbers and letters, over the selected lines', async () => {
    const { source, outline } = watched()
    const w = await mountSheet(source, { messages: ruMessages })
    await num(w, 2).trigger('click')
    await num(w, 4).trigger('click', { shiftKey: true })
    await num(w, 3).trigger('contextmenu', { button: 2, clientX: 40, clientY: 60 })
    await settle()
    expect(items().slice(-2)).toEqual(['group', 'ungroup'])
    expect(document.querySelector('[data-xs-menu="group"]')!.textContent!.trim()).toBe('Сгруппировать')
    document.querySelector<HTMLElement>('[data-xs-menu="group"]')!.click()
    await settle()
    await menuOn(head(w, 3), 'ungroup')
    expect(sent(outline)).toEqual([
      { op: 'group', sheet: 'S', rows: '2:4' },
      { op: 'ungroup', sheet: 'S', columns: 'C:C' },
    ])
    expect(w.findAll('.xs-ol-levels-rows button')).toHaveLength(4)
  })

  it('says what Excel refuses without sending it: Ungroup of lines in no group, Group past seven levels', async () => {
    const deep = book()
    deep.sheets[0]!.outline = { rows: { 3: 7 } }
    const source = memorySource(deep)
    const outline = vi.spyOn(source, 'outline')
    const w = await mountSheet(source, { messages: ruMessages })
    await menuOn(num(w, 12), 'ungroup')
    expect(w.find('.xs-message').text()).toBe(ruMessages.outlineNone)
    await menuOn(num(w, 3), 'group')
    expect(w.find('.xs-message').text()).toBe(ruMessages.outlineDeepest)
    expect(outline).not.toHaveBeenCalled()
    // with a line above it Excel deepens that line: sent
    await num(w, 2).trigger('click')
    await num(w, 3).trigger('click', { shiftKey: true })
    await menuOn(num(w, 3), 'group')
    expect(sent(outline)).toEqual([{ op: 'group', sheet: 'S', rows: '2:3' }])
  })

  it('by Excel’s keys over whole lines: Shift+Alt+→ and ←, Cmd+Shift+K and J (any layout)', async () => {
    const { source, outline } = watched()
    const w = await mountSheet(source)
    await num(w, 11).trigger('click')
    await key(w, 'ArrowRight', { shiftKey: true, altKey: true })
    await settle()
    await head(w, 2).trigger('click')
    await head(w, 3).trigger('click', { shiftKey: true })
    await key(w, 'J', { metaKey: true, shiftKey: true, code: 'KeyJ' })
    await settle()
    await key(w, 'Л', { metaKey: true, shiftKey: true, code: 'KeyK' })
    await settle()
    await num(w, 2).trigger('click')
    await key(w, 'ArrowLeft', { shiftKey: true, altKey: true })
    await settle()
    expect(sent(outline)).toEqual([
      { op: 'group', sheet: 'S', rows: '11:11' },
      { op: 'ungroup', sheet: 'S', columns: 'B:C' },
      { op: 'group', sheet: 'S', columns: 'B:C' },
      { op: 'ungroup', sheet: 'S', rows: '2:2' },
    ])
    // The keys did not move the selection on their way.
    expect(address(w)).toBe('2:2')
  })

  it('over cells the box asks rows or columns, rows first, as Excel’s does; Cancel sends nothing', async () => {
    const { source, outline } = watched()
    const w = await mountSheet(source, { messages: ruMessages })
    await goTo(w, 'B2:C3')
    await key(w, 'ArrowRight', { shiftKey: true, altKey: true })
    await settle()
    const box = document.querySelector<HTMLElement>('[data-xs-dialog]')!
    expect(box.textContent).toContain('Сгруппировать')
    expect(document.querySelector<HTMLInputElement>('[data-xs-outline-axis="rows"]')!.checked).toBe(true)
    document.querySelector<HTMLElement>('[data-xs-dialog-cancel]')!.click()
    await settle()
    expect(outline).not.toHaveBeenCalled()
    ;(w.find('.xs-scroller').element as HTMLElement).focus()
    await key(w, 'ArrowRight', { shiftKey: true, altKey: true })
    await settle()
    const columns = document.querySelector<HTMLInputElement>('[data-xs-outline-axis="columns"]')!
    columns.checked = true
    columns.dispatchEvent(new Event('change'))
    document.querySelector<HTMLElement>('[data-xs-dialog-ok]')!.click()
    await settle()
    expect(sent(outline)).toEqual([{ op: 'group', sheet: 'S', columns: 'B:C' }])
  })

  it('without the source’s outline Shift+Alt+→ extends the selection as before and the menu has no items for it', async () => {
    const plain = memorySource(book())
    delete (plain as { outline?: unknown }).outline
    const w = await mountSheet(plain)
    await num(w, 2).trigger('click')
    const e = new KeyboardEvent('keydown', { key: 'ArrowRight', shiftKey: true, altKey: true, bubbles: true, cancelable: true })
    w.find('.xs-scroller').element.dispatchEvent(e)
    await settle()
    expect(address(w)).toBe('A2:Z2')
    expect(document.querySelector('[data-xs-dialog]')).toBeNull()
    await num(w, 2).trigger('contextmenu', { button: 2 })
    await settle()
    expect(items()).not.toContain('group')
  })
})
