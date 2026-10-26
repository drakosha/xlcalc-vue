// Excel's Show Formulas (Ctrl+`): cells show their formulas and unformatted constants, columns are drawn twice as wide
// without changing the workbook, per sheet, and editing works as before.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook, type SheetViewApi, type SheetViewState } from '../src'
import { colPx, widthFromPx } from '../src/core/geometry'
import { cellText, editorKey, goTo, key, mountSheet, settle, until } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const book = (): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 1234.5, text: '1,234.50', numberFormat: '#,##0.00' },
        A2: { value: 45688, text: '1/31/2025', numberFormat: 'm/d/yyyy' },
        A3: { value: -2, text: '(2)', color: '#FF0000', numberFormat: '0;[Red](0)' },
        A4: 'Total',
        B1: { value: 12345, text: '12,345.00', formula: 'A1*10' },
        B2: { value: '', text: '', formula: 'IF(A1>0,"","x")' },
        C1: { value: true, text: 'TRUE' },
      },
      widths: { C: 20 },
    },
    { name: 'T', cells: { A1: { value: 3, formula: '1+2' } } },
  ],
})
const toggle = (w: VueWrapper) => key(w, '`', { code: 'Backquote', ctrlKey: true })
const colWidths = (w: VueWrapper) => w.findAll('col').map((c) => c.attributes('style'))
const justify = (w: VueWrapper, row: number, col: number) => w.find(`[data-cell="${row}:${col}"] .xs-in`).attributes('style') ?? ''

describe('Show Formulas', () => {
  it('Ctrl+` shows formulas with = and constants without their number formats, and Ctrl+` again shows the values', async () => {
    const states: SheetViewState[] = []
    const w = await mountSheet(memorySource(book()), { onState: (st: SheetViewState | null) => st && states.push(st) })
    await toggle(w)
    await until(() => cellText(w, 1, 1) === '1234.5')
    expect([cellText(w, 1, 2), cellText(w, 2, 2), cellText(w, 2, 1), cellText(w, 4, 1), cellText(w, 1, 3)]).toEqual(['=A1*10', '=IF(A1>0,"","x")', '45688', 'Total', 'TRUE'])
    expect(states.at(-1)!.showFormulas).toBe(true)
    await toggle(w)
    expect([cellText(w, 1, 1), cellText(w, 1, 2), cellText(w, 2, 1)]).toEqual(['1,234.50', '12,345.00', '1/31/2025'])
    expect(states.at(-1)!.showFormulas).toBeUndefined()
  })

  it('a formula is aligned left as text, a number stays right; the number format color is not applied', async () => {
    const w = await mountSheet(memorySource(book()))
    await toggle(w)
    await settle()
    expect(justify(w, 1, 2)).toContain('text-align: left')
    expect(justify(w, 1, 1)).toContain('text-align: right')
    expect(w.find('[data-cell="3:1"]').attributes('style') ?? '').not.toContain('#FF0000')
    await toggle(w)
    expect(w.find('[data-cell="3:1"]').attributes('style')).toContain('#FF0000')
  })

  it('columns are drawn twice as wide, and the stored widths do not change', async () => {
    const source = memorySource(book())
    const layout = vi.spyOn(source, 'layout')
    const w = await mountSheet(source)
    const before = colWidths(w)
    expect(before).toContain(`width: ${colPx(null)}px;`)
    await toggle(w)
    await settle()
    const after = colWidths(w)
    expect(after).toContain(`width: ${colPx(null) * 2}px;`)
    expect(after).toContain(`width: ${colPx(20) * 2}px;`)
    await toggle(w)
    expect(colWidths(w)).toEqual(before)
    expect(layout).not.toHaveBeenCalled()
  })

  it('a width dragged with the mouse is stored as half the drawn one', async () => {
    const source = memorySource(book())
    const layout = vi.spyOn(source, 'layout')
    const w = await mountSheet(source)
    await toggle(w)
    await settle()
    const grip = w.findAll('.xs-head')[1]!.find('.xs-grip')
    await grip.trigger('pointerdown', { clientX: 100, pointerId: 1 })
    await grip.trigger('pointermove', { clientX: 100 + 200 - colPx(null) * 2, pointerId: 1 })
    await grip.trigger('pointerup', { clientX: 100 + 200 - colPx(null) * 2, pointerId: 1 })
    await settle()
    expect(colWidths(w)).toContain('width: 200px;')
    expect(layout).toHaveBeenCalledWith({ sheet: 'S', columns: 'B', width: widthFromPx(100) })
    await toggle(w)
    expect(colWidths(w)).toContain('width: 100px;')
  })

  it('is a setting of each sheet: another sheet shows its values, and the sheet shown again keeps its formulas', async () => {
    const w = await mountSheet(memorySource(book()), { storageKey: 'show-formulas-test' })
    await toggle(w)
    await settle()
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    expect(cellText(w, 1, 1)).toBe('3')
    await w.findAll('.xs-tab')[0]!.trigger('click')
    await settle()
    expect(cellText(w, 1, 2)).toBe('=A1*10')
    localStorage.clear()
  })

  it('editing still edits: F2 opens the cell text, and the entry is shown as its formula', async () => {
    const w = await mountSheet(memorySource(book()))
    await toggle(w)
    await goTo(w, 'B1')
    await key(w, 'F2')
    await settle()
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('=A1*10')
    await w.find('.xs-editor').setValue('=A1*20')
    await editorKey(w, 'Enter')
    await settle()
    expect(cellText(w, 1, 2)).toBe('=A1*20')
  })

  it("the host's showFormulas: on, off, toggled, on a read-only sheet too", async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    const api = w.vm as unknown as SheetViewApi
    expect(api.showFormulas(true)).toBe(true)
    await settle()
    expect(cellText(w, 1, 2)).toBe('=A1*10')
    expect(api.showFormulas()).toBe(false)
    await settle()
    expect(cellText(w, 1, 2)).toBe('12,345.00')
    expect(api.showFormulas()).toBe(true)
  })
})
