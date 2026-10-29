// Excel's Paste Special and Format Painter over `./local` (xlcalc 0.60's copyRange with `paste` and `transpose`): the dialog
// pastes a copy as values, formats or transposed; the painter takes the selection's formats and paints the next range.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import { SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import type { SheetWindow } from '../src/contract/source'
import { cellText, clipboard, goTo, key, mountSheet, settle } from './helpers'
import { styleOf } from './local-fixture'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const open: LocalSource[] = []
afterEach(async () => {
  for (const s of open.splice(0)) await s.close()
})
/** A1 bold 1, B1 =A1*10 with a fill, A2 2; C1:D2 empty. */
function source(): LocalSource {
  const s = localSource(
    {
      locale: 'en-US',
      styles: [{}, { font: { bold: true } }, { fill: { color: '#FFFF00' } }],
      sheets: [{ name: 'S', cells: { A1: { value: 1, style: 1 }, B1: { formula: 'A1*10', style: 2 }, A2: { value: 2 } } }],
    },
    { calc: 'inline' },
  )
  open.push(s)
  return s
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
/** The sheet over a local source, once the workbook is open. */
async function mountLocal(s: LocalSource, props: Record<string, unknown> = {}): Promise<VueWrapper> {
  const w = await mountSheet(s, props)
  await settle(12)
  return w
}
async function look(s: LocalSource, cell: string) {
  const [col, row] = [cell.charCodeAt(0) - 64, Number(cell.slice(1))]
  const win = (await s.window({ sheet: 'S', row, rows: 1, col, cols: 1 })) as SheetWindow
  return { text: win.text[0]![0], formula: win.formulas[cell], style: styleOf(win, cell) }
}
const dialog = (w: VueWrapper) => w.find('[data-xs-paste-special]')
async function pasteSpecial(w: VueWrapper, kind: string, transpose = false, more: { operation?: string; skipBlanks?: boolean } = {}): Promise<void> {
  expect(api(w).pasteSpecial()).toBe(true)
  await settle()
  await w.find(`[data-xs-paste-kind="${kind}"]`).setValue(true)
  if (transpose) await w.find('[data-xs-paste-transpose]').setValue(true)
  if (more.operation) await w.find(`[data-xs-paste-operation="${more.operation}"]`).setValue(true)
  if (more.skipBlanks) await w.find('[data-xs-paste-skip]').setValue(true)
  await w.find('[data-xs-paste-special] [data-xs-dialog-ok]').trigger('click')
  await settle(12)
}
/** A sheet over `cells` (and `extra` of its sheet), mounted once the workbook is open. */
async function sheetOf(cells: Record<string, unknown>, extra: Record<string, unknown> = {}): Promise<{ s: LocalSource; w: VueWrapper }> {
  const s = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells, ...extra }] } as never, { calc: 'inline' })
  open.push(s)
  return { s, w: await mountLocal(s) }
}

describe('Paste Special', () => {
  it('pastes a copy as values (no formula, no formats) or as formats (no contents), each a single undo step', async () => {
    const s = source()
    const w = await mountLocal(s)
    await goTo(w, 'A1:B1')
    clipboard(w, 'copy')
    await goTo(w, 'C3')
    await pasteSpecial(w, 'values')
    expect(dialog(w).exists()).toBe(false)
    expect(await look(s, 'D3')).toEqual({ text: '10', formula: undefined, style: undefined })
    expect((await look(s, 'C3')).style).toBeUndefined()
    await goTo(w, 'C5')
    await pasteSpecial(w, 'formats')
    expect(await look(s, 'C5')).toEqual({ text: '', formula: undefined, style: { font: { bold: true } } })
    expect((await look(s, 'D5')).style).toEqual({ fill: { color: '#FFFF00' } })
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    expect((await look(s, 'C5')).style).toBeUndefined()
  })

  it('transposes: a row of the copy becomes a column', async () => {
    const s = source()
    const w = await mountLocal(s)
    await goTo(w, 'A1:A2')
    clipboard(w, 'copy')
    await goTo(w, 'C1')
    await pasteSpecial(w, 'all', true)
    expect([cellText(w, 1, 3), cellText(w, 1, 4)]).toEqual(['1', '2'])
  })

  it("offers Excel's Paste choices, its Operation group, Skip blanks and Transpose; All and None at first", async () => {
    const { w } = await sheetOf({ A1: { value: 1 } })
    await goTo(w, 'A1')
    clipboard(w, 'copy')
    api(w).pasteSpecial()
    await settle()
    const values = (sel: string, attr: string) => w.findAll(sel).map((x) => x.attributes(attr))
    expect(values('[data-xs-paste-kind]', 'data-xs-paste-kind')).toEqual(['all', 'formulas', 'values', 'formats', 'notes', 'validation', 'allExceptBorders', 'columnWidths', 'formulasAndNumberFormats', 'valuesAndNumberFormats', 'allMergingConditionalFormats'])
    expect(values('[data-xs-paste-operation]', 'data-xs-paste-operation')).toEqual(['none', 'add', 'subtract', 'multiply', 'divide'])
    const checked = (sel: string) => (w.find(sel).element as HTMLInputElement).checked
    expect([checked('[data-xs-paste-kind="all"]'), checked('[data-xs-paste-operation="none"]'), checked('[data-xs-paste-skip]'), checked('[data-xs-paste-transpose]')]).toEqual([true, true, false, false])
    // Excel for Windows' labels in its screen order, without All using Source theme and Paste Link
    expect(w.findAll('[data-xs-paste-special] label').map((x) => x.text())).toEqual([
      'All', 'Formulas', 'Values', 'Formats', 'Comments', 'Validation',
      'All except borders', 'Column widths', 'Formulas and number formats', 'Values and number formats', 'All merging conditional formats',
      'None', 'Add', 'Subtract', 'Multiply', 'Divide', 'Skip blanks', 'Transpose',
    ])
  })

  it("combines the copy's numbers with the cells under it (Add, Multiply); a formula on either side makes a formula of both", async () => {
    const { s, w } = await sheetOf({ A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'A1*2' }, C1: { value: 10 }, C2: { value: 10 }, D1: { value: 3 } })
    await goTo(w, 'A1:A2')
    clipboard(w, 'copy')
    await goTo(w, 'C1')
    await pasteSpecial(w, 'values', false, { operation: 'add' })
    expect([(await look(s, 'C1')).text, (await look(s, 'C2')).text]).toEqual(['11', '12'])
    await goTo(w, 'B1')
    clipboard(w, 'copy')
    await goTo(w, 'D1')
    await pasteSpecial(w, 'formulas', false, { operation: 'multiply' })
    expect(await look(s, 'D1')).toMatchObject({ formula: '3*(C1*2)', text: '66' })
  })

  it("Skip Blanks keeps the cells under the copy's blank cells", async () => {
    const { s, w } = await sheetOf({ A1: { value: 1 }, A3: { value: 3 }, C1: { value: 7 }, C2: { value: 8 }, C3: { value: 9 } })
    await goTo(w, 'A1:A3')
    clipboard(w, 'copy')
    await goTo(w, 'C1')
    await pasteSpecial(w, 'all', false, { skipBlanks: true })
    expect([(await look(s, 'C1')).text, (await look(s, 'C2')).text, (await look(s, 'C3')).text]).toEqual(['1', '8', '3'])
  })

  it('pastes column widths, notes, and values with their number formats', async () => {
    const { s, w } = await sheetOf({ A1: { value: 0.5, numberFormat: '0%', note: { text: 'checked' } }, B1: { formula: 'A1*2' } }, { columns: { A: { width: 20 } } })
    await goTo(w, 'A1')
    clipboard(w, 'copy')
    await goTo(w, 'F1')
    await pasteSpecial(w, 'columnWidths')
    const win = (await s.window({ sheet: 'S', row: 1, rows: 1, col: 6, cols: 1 })) as SheetWindow
    expect(win.widths[0]).toBe(20)
    expect(win.text[0]![0]).toBe('')
    await goTo(w, 'C3')
    await pasteSpecial(w, 'notes')
    const noted = (await s.window({ sheet: 'S', row: 3, rows: 1, col: 3, cols: 1 })) as SheetWindow
    expect([noted.notes?.C3?.text, noted.text[0]![0]]).toEqual(['checked', ''])
    await goTo(w, 'A1:B1')
    clipboard(w, 'copy')
    await goTo(w, 'C5')
    await pasteSpecial(w, 'valuesAndNumberFormats')
    expect([(await look(s, 'C5')).text, (await look(s, 'D5')).text, (await look(s, 'D5')).formula]).toEqual(['50%', '1', undefined])
  })

  it("shows the library's refusal of an operation onto a spilled array in the message line, and pastes nothing", async () => {
    const { s, w } = await sheetOf({ A1: { value: 1 }, B1: { formula: 'SEQUENCE(3)' } })
    await goTo(w, 'A1')
    clipboard(w, 'copy')
    await goTo(w, 'B2')
    await pasteSpecial(w, 'all', false, { operation: 'add' })
    expect(w.find('.xs-message').text()).not.toBe('')
    expect((await look(s, 'B2')).text).toBe('2')
    const out = await api(w).structure({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B2' }, operation: 'add' })
    expect(out).toMatchObject({ ok: false, code: 'array' })
  })

  it('opens with Ctrl+Alt+V and from the cell menu after a copy; not after a cut or without a copy; Cancel pastes nothing', async () => {
    const s = source()
    const w = await mountLocal(s)
    expect(api(w).pasteSpecial()).toBe(false)
    await goTo(w, 'A1')
    clipboard(w, 'cut')
    expect(api(w).pasteSpecial()).toBe(false)
    await key(w, 'Escape')
    clipboard(w, 'copy')
    await goTo(w, 'C3')
    await w.find('.xs-scroller').trigger('keydown', { key: 'v', code: 'KeyV', ctrlKey: true, altKey: true })
    await settle()
    expect(dialog(w).exists()).toBe(true)
    await w.find('[data-xs-paste-special] [data-xs-dialog-cancel]').trigger('click')
    await settle(12)
    expect(dialog(w).exists()).toBe(false)
    expect(cellText(w, 3, 3)).toBe('')
  })
})

describe('Paste Special on a protected sheet', () => {
  /** A1 locked, B1 unlocked with a note and its own width; protected with `protection`. */
  async function form(protection: Record<string, boolean>) {
    const s = localSource({ locale: 'en-US', styles: [{}, { protection: { locked: false } }], sheets: [{ name: 'S', cells: { A1: { value: 'a' }, B1: { value: 'b', style: 1, note: { text: 'n' } } }, columns: { B: { width: 30 } }, protection }] } as never, { calc: 'inline' })
    open.push(s)
    const w = await mountLocal(s)
    await goTo(w, 'B1')
    clipboard(w, 'copy')
    await settle()
    await goTo(w, 'A1')
    return { s, w }
  }
  const alert = () => document.querySelector('[data-xs-alert="notice"] [data-xs-alert-message]')?.textContent ?? null
  it('pastes notes with Edit objects and column widths with Format columns over a locked cell, as the source allows them', async () => {
    const { s, w } = await form({ objects: true, formatColumns: true })
    await pasteSpecial(w, 'notes')
    await pasteSpecial(w, 'columnWidths')
    expect(alert()).toBeNull()
    const book = await s.book()
    expect(book.sheets[0]!.cells.A1!.note).toMatchObject({ text: 'n' })
    expect(book.sheets[0]!.columns?.A?.width).toBe(30)
  })
  it("refuses them with Excel's alert before sending without their boxes, and Validation always", async () => {
    for (const kind of ['notes', 'columnWidths', 'validation']) {
      const { s, w } = await form({ objects: kind !== 'notes', formatColumns: kind !== 'columnWidths' })
      const structure = vi.spyOn(s, 'structure')
      await pasteSpecial(w, kind)
      expect([kind, alert()?.startsWith('The cell or chart'), structure.mock.calls.length]).toEqual([kind, true, 0])
      document.querySelector<HTMLButtonElement>('[data-xs-alert-button="ok"]')!.click()
      await settle()
      w.unmount()
    }
  })
})

describe('Format Painter', () => {
  const press = async (w: VueWrapper, cell: string) => {
    await goTo(w, cell)
    const el = w.find('.xs-scroller').element
    el.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }))
    el.dispatchEvent(new PointerEvent('pointerup', { bubbles: true, button: 0 }))
    // the painting waits a task for the selection a header's click makes
    await new Promise((done) => setTimeout(done, 0))
    await settle(12)
  }

  it("paints the brush's formats onto the next range once; with keep until Esc; Esc puts it down", async () => {
    const s = source()
    const states: SheetViewState[] = []
    const w = await mountLocal(s, { onState: (st: SheetViewState | null) => st && states.push(st) })
    await goTo(w, 'A1')
    expect(states.at(-1)!.canFormatPainter).toBe(true)
    expect(api(w).formatPainter()).toBe(true)
    await settle()
    expect(states.at(-1)!.painting).toBe('once')
    expect(w.find('[data-xs-ants]').exists()).toBe(true)
    await press(w, 'C3')
    expect((await look(s, 'C3')).style).toEqual({ font: { bold: true } })
    expect(states.at(-1)!.painting).toBeUndefined()
    await goTo(w, 'B1')
    expect(api(w).formatPainter('keep')).toBe(true)
    await press(w, 'C4')
    await press(w, 'C5')
    expect([(await look(s, 'C4')).style, (await look(s, 'C5')).style]).toEqual([{ fill: { color: '#FFFF00' } }, { fill: { color: '#FFFF00' } }])
    expect(states.at(-1)!.painting).toBe('keep')
    await key(w, 'Escape')
    await settle()
    expect(states.at(-1)!.painting).toBeUndefined()
    await press(w, 'C6')
    expect((await look(s, 'C6')).style).toBeUndefined()
  })

  it("a larger font painted onto a row without its own height fits the row, as Excel does (xlcalc 0.60.1's fitted)", async () => {
    const s = localSource({ locale: 'en-US', styles: [{}, { font: { size: 20, bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 'Big', style: 1 }, C5: { value: 'x' } } }] }, { calc: 'inline' })
    open.push(s)
    const w = await mountLocal(s)
    await goTo(w, 'A1')
    api(w).formatPainter()
    await press(w, 'C5')
    const win = (await s.window({ sheet: 'S', row: 5, rows: 1, col: 3, cols: 1 })) as SheetWindow
    expect(win.rowHeights[5]).toBe(26)
  })

  it('is not available on a read-only sheet', async () => {
    const w = await mountLocal(source(), { readOnly: true })
    await goTo(w, 'A1')
    expect(api(w).formatPainter()).toBe(false)
  })

  it("the toolbar's button: a click takes the brush or puts it down, a double click keeps it; pressed while held", async () => {
    const s = source()
    const sheet = ref<SheetViewApi | null>(null)
    const state = ref<SheetViewState | null>(null)
    const Host = defineComponent(() => () => [
      h(SheetToolbar, { source: s, sheet: sheet.value, state: state.value }),
      h(SheetView, { ref: sheet, source: s, style: 'height: 400px', onState: (st: SheetViewState | null) => (state.value = st) }),
    ])
    const w = mount(Host, { attachTo: document.body })
    await settle(12)
    await goTo(w, 'A1')
    const button = () => w.find('[data-tb="format-painter"]')
    expect(button().attributes('disabled')).toBeUndefined()
    await button().trigger('click')
    await settle()
    expect([state.value?.painting, button().attributes('aria-pressed')]).toEqual(['once', 'true'])
    await button().trigger('click')
    await settle()
    expect(state.value?.painting).toBeUndefined()
    await button().trigger('dblclick')
    await settle()
    expect(state.value?.painting).toBe('keep')
  })
})
