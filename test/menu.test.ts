// The structure menu: on row numbers and column letters it inserts and deletes the selected lines, sets their size in
// Excel's units, hides and shows them; on the tabs it inserts, renames, moves, deletes, hides and shows sheets. A right
// click opens it, a finger's long press too; a refusal shows the source's message.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, ruMessages, type MemoryBook } from '../src'
import { widthFromChars } from '../src/core/geometry'
import { localSource } from '../src/local'
import type { EditOutcome, SheetOp } from '../src/source'
import { cellText, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const OK: EditOutcome = { ok: true, changes: [] }
const book = (): MemoryBook => ({
  sheets: [
    { name: 'S', cells: { A1: 1, A2: 2, A3: { value: 3, formula: 'SUM(A1:A2)' }, D6: 'x' } },
    { name: 'T', cells: { A1: 't' } },
    { name: 'U', cells: {} },
  ],
})
/** A memory book with a structure method the test watches. */
function withStructure(answer: (op: SheetOp) => EditOutcome = () => OK) {
  const source = memorySource(book())
  const structure = vi.fn(async (op: SheetOp) => answer(op))
  source.structure = structure
  const layout = vi.spyOn(source, 'layout')
  return { source, structure, layout }
}
const num = (w: VueWrapper, row: number) => w.findAll('.xs-num').find((n) => n.text() === String(row))!
const head = (w: VueWrapper, col: number) => w.findAll('.xs-head')[col - 1]!
const items = () => [...document.querySelectorAll<HTMLElement>('[data-xs-menu]')].map((b) => b.dataset.xsMenu)
const pick = async (what: string) => {
  document.querySelector<HTMLElement>(`[data-xs-menu="${what}"]`)!.click()
  await settle()
}
async function context(el: { trigger: (e: string, o?: object) => Promise<void> }): Promise<void> {
  await el.trigger('contextmenu', { button: 2, clientX: 40, clientY: 60 })
  await settle()
}
const dialog = () => document.querySelector<HTMLElement>('[data-xs-dialog]')
async function answer(text: string): Promise<void> {
  const input = document.querySelector<HTMLInputElement>('[data-xs-dialog-input]')!
  input.value = text
  input.dispatchEvent(new Event('input'))
  document.querySelector<HTMLElement>('[data-xs-dialog-ok]')!.click()
  await settle()
}

describe('the menu on row numbers and column letters', () => {
  it('inserts and deletes the selected rows, or the row clicked outside the selection, as Excel', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await context(num(w, 3))
    expect(items()).toEqual(['insert', 'delete', 'size', 'hide', 'group', 'ungroup'])
    await pick('insert')
    await num(w, 2).trigger('click')
    await num(w, 4).trigger('click', { shiftKey: true })
    await context(num(w, 3))
    await pick('insert')
    await context(num(w, 6))
    await pick('delete')
    expect(structure.mock.calls.map((c) => c[0])).toEqual([
      { op: 'insertRows', sheet: 'S', row: 3, count: 1 },
      { op: 'insertRows', sheet: 'S', row: 2, count: 3 },
      { op: 'deleteRows', sheet: 'S', row: 6, count: 1 },
    ])
    expect(w.find('.xs-address').element).toHaveProperty('value', '6:6')
  })

  it('inserts and deletes columns by their letters', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await context(head(w, 3))
    await pick('insert')
    await head(w, 2).trigger('click')
    await head(w, 4).trigger('click', { shiftKey: true })
    await context(head(w, 2))
    await pick('delete')
    expect(structure.mock.calls.map((c) => c[0])).toEqual([
      { op: 'insertColumns', sheet: 'S', column: 'C', count: 1 },
      { op: 'deleteColumns', sheet: 'S', column: 'B', count: 3 },
    ])
  })

  it('hides lines, and shows them on a selection over hidden ones', async () => {
    const { source, layout } = withStructure()
    const w = await mountSheet(source)
    await head(w, 2).trigger('click')
    await head(w, 3).trigger('click', { shiftKey: true })
    await context(head(w, 2))
    await pick('hide')
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', columns: 'B:C', hidden: true })
    expect(w.findAll('.xs-head').map((h) => h.text()).slice(0, 3)).toEqual(['A', 'D', 'E'])
    await head(w, 1).trigger('click')
    await w.findAll('.xs-head')[1]!.trigger('click', { shiftKey: true })
    await context(head(w, 1))
    expect(items()).toEqual(['insert', 'delete', 'size', 'hide', 'unhide', 'group', 'ungroup'])
    await pick('unhide')
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', columns: 'A:D', hidden: false })
    expect(w.findAll('.xs-head').map((h) => h.text()).slice(0, 3)).toEqual(['A', 'B', 'C'])
    await context(num(w, 2))
    await pick('hide')
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', rows: '2', hidden: true })
    expect(w.findAll('.xs-num').map((n) => n.text()).slice(0, 3)).toEqual(['1', '3', '4'])
  })

  it('sets a column width in characters and a row height in points, as Excel’s boxes; 0 hides; out of range is refused', async () => {
    const { source, layout } = withStructure()
    const w = await mountSheet(source)
    await context(head(w, 2))
    await pick('size')
    expect(document.querySelector<HTMLInputElement>('[data-xs-dialog-input]')!.value).toBe('8.14')
    await answer('300')
    expect(dialog()!.textContent).toContain('Column width must be between 0 and 255')
    await answer('20')
    expect(dialog()).toBeNull()
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', columns: 'B', width: widthFromChars(20) })
    await context(num(w, 2))
    await pick('size')
    expect(document.querySelector<HTMLInputElement>('[data-xs-dialog-input]')!.value).toBe('15')
    await answer('30')
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', rows: '2', height: 30 })
    await context(num(w, 4))
    await pick('size')
    await answer('0')
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', rows: '4', hidden: true })
  })

  it('reads and writes the sizes with the book’s decimal comma', async () => {
    const { source, layout } = withStructure()
    const w = await mountSheet(source, { locale: 'ru-RU', messages: ruMessages })
    await context(head(w, 3))
    expect(document.querySelector('[data-xs-menu="size"]')!.textContent!.trim()).toBe(ruMessages.menuColumnWidth)
    await pick('size')
    expect(document.querySelector<HTMLInputElement>('[data-xs-dialog-input]')!.value).toBe('8,14')
    await answer('12,5')
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', columns: 'C', width: widthFromChars(12.5) })
  })

  it('shows a refusal by the source’s message', async () => {
    const { source } = withStructure(() => ({ ok: false, reason: 'refused', message: 'Нельзя сдвинуть непустые ячейки за край листа' }))
    const w = await mountSheet(source)
    await context(num(w, 2))
    await pick('insert')
    expect(w.find('.xs-message').text()).toBe('Нельзя сдвинуть непустые ячейки за край листа')
  })

  it('has no structure items without structure, no menu without layout and outline too, none on a read-only sheet', async () => {
    const plain = memorySource(book())
    const w = await mountSheet(plain)
    await context(num(w, 2))
    expect(items()).toEqual(['size', 'hide', 'group', 'ungroup'])
    document.body.innerHTML = ''
    delete (plain as { layout?: unknown }).layout
    delete (plain as { outline?: unknown }).outline
    const bare = await mountSheet(plain)
    const e = new MouseEvent('contextmenu', { bubbles: true, cancelable: true })
    num(bare, 2).element.dispatchEvent(e)
    await settle()
    expect(e.defaultPrevented).toBe(false)
    expect(items()).toEqual([])
    const ro = await mountSheet(withStructure().source, { readOnly: true })
    await context(num(ro, 2))
    expect(items()).toEqual([])
  })

  it('closes on Escape and on a press outside', async () => {
    const w = await mountSheet(withStructure().source)
    await context(num(w, 2))
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    await settle()
    expect(items()).toEqual([])
    await context(num(w, 2))
    w.find('[data-cell="5:3"]').element.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    await settle()
    expect(items()).toEqual([])
  })
})

describe('a finger’s long press', () => {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms))
  const touch = (x = 20, y = 30) => ({ pointerType: 'touch', isPrimary: true, button: 0, clientX: x, clientY: y })

  it('opens the menu after half a second; a move (a scroll) or a lift before that does not', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await num(w, 3).trigger('pointerdown', touch())
    await num(w, 3).trigger('pointermove', touch(20, 60))
    await wait(560)
    await settle()
    expect(items()).toEqual([])
    await num(w, 3).trigger('pointerdown', touch())
    await num(w, 3).trigger('pointerup', touch())
    await wait(560)
    expect(items()).toEqual([])
    await num(w, 3).trigger('pointerdown', touch())
    await num(w, 3).trigger('pointermove', touch(23, 32))
    await wait(560)
    await settle()
    expect(items()).toEqual(['insert', 'delete', 'size', 'hide', 'group', 'ungroup'])
    // The lift and the browser's own context menu after it belong to the press: the selection stays, the menu too.
    await num(w, 3).trigger('pointerup', touch())
    await num(w, 3).trigger('contextmenu')
    await num(w, 3).trigger('click')
    await settle()
    expect(items()).toEqual(['insert', 'delete', 'size', 'hide', 'group', 'ungroup'])
    await pick('insert')
    expect(structure).toHaveBeenCalledWith({ op: 'insertRows', sheet: 'S', row: 3, count: 1 })
  })

  it('opens the tab’s menu too', async () => {
    const w = await mountSheet(withStructure().source)
    await w.findAll('.xs-tab')[1]!.trigger('pointerdown', touch())
    await wait(560)
    await settle()
    expect(w.find('.xs-tab-on').text()).toBe('T')
    expect(items()).toEqual(['insertSheet', 'deleteSheet', 'rename', 'moveLeft', 'moveRight', 'hideSheet'])
  })
})

describe('the menu on the tabs', () => {
  it('inserts a sheet before the one clicked, shows the next after deleting an empty one, moves by the shown tabs', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    const tab = (i: number) => w.findAll('.xs-tab')[i]!
    await context(tab(1))
    expect(w.find('.xs-tab-on').text()).toBe('T')
    await pick('insertSheet')
    await context(tab(2))
    await pick('moveLeft')
    await context(tab(0))
    expect(document.querySelector('[data-xs-menu="moveLeft"]')!.hasAttribute('disabled')).toBe(true)
    await pick('moveRight')
    await context(tab(2))
    await pick('deleteSheet')
    expect(structure.mock.calls.map((c) => c[0])).toEqual([
      { op: 'insertSheet', name: 'Sheet4', before: 'T' },
      { op: 'moveSheet', sheet: 'U', before: 'T' },
      { op: 'moveSheet', sheet: 'S', before: 'U' },
      { op: 'deleteSheet', sheet: 'U' },
    ])
  })

  it('asks before deleting a sheet with data, as Excel does', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await context(w.findAll('.xs-tab')[1]!)
    await pick('deleteSheet')
    expect(structure).not.toHaveBeenCalled()
    expect(dialog()!.textContent).toContain('“T”')
    document.querySelector<HTMLElement>('[data-xs-dialog-cancel]')!.click()
    await settle()
    expect(dialog()).toBeNull()
    await context(w.findAll('.xs-tab')[1]!)
    await pick('deleteSheet')
    document.querySelector<HTMLElement>('[data-xs-dialog-ok]')!.click()
    await settle()
    expect(structure).toHaveBeenCalledWith({ op: 'deleteSheet', sheet: 'T' })
  })

  it('renames on the tab itself: from the menu or a double click, Enter keeps, Esc drops; a refusal keeps the field', async () => {
    let refuse = true
    const { source, structure } = withStructure(() => (refuse ? { ok: false, reason: 'refused', message: 'Лист «T» уже есть' } : OK))
    const w = await mountSheet(source)
    await context(w.findAll('.xs-tab')[0]!)
    await pick('rename')
    const field = () => w.find<HTMLInputElement>('.xs-tab-input')
    expect(field().element.value).toBe('S')
    expect(document.activeElement).toBe(field().element)
    await field().setValue('T')
    await field().trigger('keydown', { key: 'Enter' })
    await settle()
    expect(w.find('.xs-message').text()).toBe('Лист «T» уже есть')
    expect(field().exists()).toBe(true)
    await field().trigger('keydown', { key: 'Escape' })
    expect(field().exists()).toBe(false)
    refuse = false
    await w.findAll('.xs-tab')[0]!.trigger('dblclick')
    await settle()
    await field().setValue('Итоги')
    await field().trigger('keydown', { key: 'Enter' })
    await settle()
    expect(structure.mock.calls.map((c) => c[0])).toEqual([
      { op: 'renameSheet', sheet: 'S', name: 'T' },
      { op: 'renameSheet', sheet: 'S', name: 'Итоги' },
    ])
    expect(field().exists()).toBe(false)
  })

  it('hides a sheet; "Unhide…" lists the hidden ones and shows the one picked', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: {} }, { name: 'Скрытый', hidden: true, cells: {} }, { name: 'Ещё', hidden: true, cells: {} }] })
    const structure = vi.fn(async (_op: SheetOp) => OK)
    source.structure = structure
    const w = await mountSheet(source, { messages: ruMessages })
    await context(w.findAll('.xs-tab')[0]!)
    expect(items()).toEqual(['insertSheet', 'deleteSheet', 'rename', 'moveLeft', 'moveRight', 'hideSheet', 'unhideSheet'])
    await pick('hideSheet')
    await context(w.findAll('.xs-tab')[0]!)
    await pick('unhideSheet')
    expect([...document.querySelectorAll<HTMLElement>('[data-xs-unhide]')].map((b) => b.textContent!.trim())).toEqual(['Скрытый', 'Ещё'])
    document.querySelector<HTMLElement>('[data-xs-unhide="Ещё"]')!.click()
    await settle()
    expect(structure.mock.calls.map((c) => c[0])).toEqual([
      { op: 'hideSheet', sheet: 'S' },
      { op: 'unhideSheet', sheet: 'Ещё' },
    ])
  })

  it('does not offer a very hidden sheet to show again, as Excel does not', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: {} }, { name: 'Скрытый', hidden: true, cells: {} }, { name: 'Тайный', hidden: 'very', cells: {} }] })
    source.structure = vi.fn(async (_op: SheetOp) => OK)
    const w = await mountSheet(source, { messages: ruMessages })
    expect(w.findAll('.xs-tab').map((t) => t.text())).toEqual(['S'])
    await context(w.findAll('.xs-tab')[0]!)
    await pick('unhideSheet')
    expect([...document.querySelectorAll<HTMLElement>('[data-xs-unhide]')].map((b) => b.textContent!.trim())).toEqual(['Скрытый'])
  })
})

describe('with the local source', () => {
  it('a row inserted from the menu moves the formula under it, as Excel', async () => {
    const source = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { formula: 'SUM(A1:A2)' } } }] }, { calc: 'inline' })
    const states: { formula?: string }[] = []
    const w = await mountSheet(source, { onState: (st: { formula?: string } | null) => st && states.push(st) })
    await settle(12)
    expect(cellText(w, 3, 1)).toBe('3')
    await context(num(w, 2))
    await pick('insert')
    await settle(12)
    expect([cellText(w, 3, 1), cellText(w, 4, 1)]).toEqual(['2', '3'])
    await w.find('[data-cell="4:1"]').trigger('pointerdown', { button: 0 })
    await settle()
    expect(states.at(-1)!.formula).toBe('SUM(A1:A3)')
    await source.close()
  })
})
