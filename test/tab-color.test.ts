// Excel's Tab Color: a colored tab is filled with its color, the active one underlined in it, and the tab menu's Tab Color
// sets or removes the color through the source's `tabColor`, a single call.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ruMessages } from '../src'
import type { EditOutcome, SheetInfo, SheetOp } from '../src/contract/source'
import { darkColour } from '../src/core/sheet-look'
import { fakeSource, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

/** A workbook whose sheets the test holds, with `structure` (the tab menu) and `tabColor`, which colors a tab. */
function book(sheets: SheetInfo[], withColor = true) {
  const f = fakeSource(sheets)
  f.source.sheets = async () => sheets.map((x) => ({ ...x }))
  f.source.structure = async (_op: SheetOp): Promise<EditOutcome> => ({ ok: true, changes: [] })
  const tabColor = vi.fn(async (sheet: string, color: string | null): Promise<EditOutcome> => {
    const at = sheets.findIndex((x) => x.name === sheet)
    const { tabColor: _, ...rest } = sheets[at]!
    sheets[at] = color ? { ...rest, tabColor: color } : rest
    return { ok: true, changes: [] }
  })
  if (withColor) f.source.tabColor = tabColor
  return { source: f.source, tabColor }
}

const tab = (w: VueWrapper, name: string) => w.findAll('.xs-tab').find((t) => t.text() === name)!
async function openTabMenu(w: VueWrapper, name: string): Promise<void> {
  await tab(w, name).trigger('contextmenu', { clientX: 10, clientY: 10 })
  await settle()
}

describe('tab color', () => {
  it('fills another tab with its color and underlines the active one in it, as Excel draws them', async () => {
    const b = book([
      { name: 'A', maxRow: 2, maxCol: 2, tabColor: '#FF0000' },
      { name: 'B', maxRow: 2, maxCol: 2, tabColor: '#FFFF00' },
      { name: 'C', maxRow: 2, maxCol: 2 },
    ])
    const w = await mountSheet(b.source)
    const a = tab(w, 'A')
    expect(a.classes()).toEqual(expect.arrayContaining(['xs-tab-on', 'xs-tab-colored']))
    expect(a.attributes('style')).toContain('--xs-tab-color: #FF0000')
    // a dark color takes light text, a light one dark text
    expect(a.classes()).toContain('xs-tab-dark')
    expect(tab(w, 'B').classes()).toContain('xs-tab-colored')
    expect(tab(w, 'B').classes()).not.toContain('xs-tab-dark')
    expect(tab(w, 'C').classes()).not.toContain('xs-tab-colored')
    expect(tab(w, 'C').attributes('style')).toBeUndefined()
  })

  it("tells a dark color from a light one by its luminance, as Excel picks a tab's text", () => {
    expect(darkColour('#000000')).toBe(true)
    expect(darkColour('#FF0000')).toBe(true)
    expect(darkColour('#0070C0')).toBe(true)
    expect(darkColour('#FFFF00')).toBe(false)
    expect(darkColour('#92D050')).toBe(false)
    expect(darkColour('#FFFFFF')).toBe(false)
  })

  it("the tab menu's Tab Color sets a color from Excel's palette and No Color removes it, each a single call", async () => {
    const b = book([
      { name: 'A', maxRow: 2, maxCol: 2 },
      { name: 'B', maxRow: 2, maxCol: 2 },
    ])
    const w = await mountSheet(b.source)
    await openTabMenu(w, 'B')
    const ids = w.findAll('[data-xs-menu]').map((x) => x.attributes('data-xs-menu'))
    // after Move Right and before Hide, as Excel lists it
    expect(ids.indexOf('tabColor')).toBe(ids.indexOf('moveRight') + 1)
    expect(ids.indexOf('hideSheet')).toBeGreaterThan(ids.indexOf('tabColor'))
    await w.find('[data-xs-menu="tabColor"]').trigger('click')
    await settle()
    expect(w.find('[data-xs-tab-colors]').exists()).toBe(true)
    await w.find('[data-xs-tab-colors] [data-xs-colour="#FF0000"]').trigger('click')
    await settle()
    expect(b.tabColor.mock.calls).toEqual([['B', '#FF0000']])
    expect(w.find('[data-xs-tab-colors]').exists()).toBe(false)
    expect(tab(w, 'B').attributes('style')).toContain('--xs-tab-color: #FF0000')

    await openTabMenu(w, 'B')
    await w.find('[data-xs-menu="tabColor"]').trigger('click')
    await settle()
    // the tab's color shows pressed, No Color not
    expect(w.find('[data-xs-tab-colors] [data-xs-colour="#FF0000"]').attributes('aria-pressed')).toBe('true')
    expect(w.find('[data-xs-tab-color-none]').attributes('aria-pressed')).toBe('false')
    await w.find('[data-xs-tab-color-none]').trigger('click')
    await settle()
    expect(b.tabColor.mock.calls[1]).toEqual(['B', null])
    expect(tab(w, 'B').classes()).not.toContain('xs-tab-colored')
  })

  it("names the item and the dialog in the interface's language", async () => {
    const b = book([{ name: 'A', maxRow: 2, maxCol: 2 }])
    const w = await mountSheet(b.source, { messages: ruMessages })
    await openTabMenu(w, 'A')
    expect(w.find('[data-xs-menu="tabColor"]').text()).toBe('Цвет ярлычка')
    await w.find('[data-xs-menu="tabColor"]').trigger('click')
    await settle()
    expect(w.find('[data-xs-tab-color-none]').text()).toBe('Нет цвета')
  })

  it('is not offered without the source’s tabColor, or on a read-only sheet', async () => {
    const sheets = [{ name: 'A', maxRow: 2, maxCol: 2 }]
    const w = await mountSheet(book(sheets, false).source)
    await openTabMenu(w, 'A')
    expect(w.find('[data-xs-menu="tabColor"]').exists()).toBe(false)
    expect(w.find('[data-xs-menu="rename"]').exists()).toBe(true)
    const ro = await mountSheet(book(sheets).source, { readOnly: true })
    await openTabMenu(ro, 'A')
    expect(ro.find('[data-xs-menu="tabColor"]').exists()).toBe(false)
  })
})
