// The workbook's default font (SheetInfo.font): the grid's cells without a font of their own are drawn in it; without it, in --xs-font as before.
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource } from '../src'
import { mountSheet } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (font?: { name: string; size: number }) => ({
  styles: [{ font: { name: 'Arial', size: 9 } }],
  sheets: [{ name: 'S', ...(font ? { font } : {}), cells: { A1: 'plain', B1: { value: 'own', style: 0 } } }],
})

describe('the default font', () => {
  it('is set on the grid for cells without a font of their own; a cell with its own font keeps it', async () => {
    const w = await mountSheet(memorySource(book({ name: 'Calibri', size: 14 })))
    const table = w.find('.xs-table').element as HTMLElement
    expect(table.style.fontFamily).toContain('Calibri')
    // 11 pt is the grid's 12.5 px, the rest in proportion
    expect(table.style.fontSize).toBe('15.9px')
    expect(w.find('[data-cell="1:1"]').attributes('style') ?? '').not.toContain('font-family')
    const own = w.find('[data-cell="1:2"]').element as HTMLElement
    expect(own.style.fontFamily).toContain('Arial')
    expect(own.style.fontSize).toBe('10.2px')
  })

  it('without one the grid sets no font, and cells take --xs-font as before', async () => {
    const w = await mountSheet(memorySource(book()))
    const table = w.find('.xs-table').element as HTMLElement
    expect(table.style.fontFamily).toBe('')
    expect(table.style.fontSize).toBe('')
  })
})
