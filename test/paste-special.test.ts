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
async function pasteSpecial(w: VueWrapper, kind: string, transpose = false): Promise<void> {
  expect(api(w).pasteSpecial()).toBe(true)
  await settle()
  await w.find(`[data-xs-paste-kind="${kind}"]`).setValue(true)
  if (transpose) await w.find('[data-xs-paste-transpose]').setValue(true)
  await w.find('[data-xs-paste-special] [data-xs-dialog-ok]').trigger('click')
  await settle(12)
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
