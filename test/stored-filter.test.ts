import { afterEach, describe, expect, it } from 'vitest'
import type { FilterCriteria } from 'xlcalc/contract'
import type { SheetViewApi } from '../src'
import { openXlsx, saveXlsx } from '../src/local'
import { closeAll, source } from './local-fixture'
import { mountSheet, settle, until } from './helpers'

afterEach(async () => { await closeAll(); document.body.innerHTML = ''; localStorage.clear() })
const criteria = (): FilterCriteria => ({ columns: [{ column: 1, filter: { type: 'values', values: ['a'] } }] })
const book = (saved = criteria()) => ({ sheets: [{ name: 'S', cells: { A1: { value: 'Key' }, B1: { value: 'Status' }, A2: { value: 'one' }, B2: { value: 'a' }, A3: { value: 'two' }, B3: { value: 'b' } }, autoFilter: { range: 'A1:B3', criteria: saved }, rows: { '3': { hidden: true } }, filtered: true }] })

describe('stored AutoFilter criteria', () => {
  it('exposes criteria, restores the checkboxes and marks the filtered column', async () => {
    const s = source(book())
    expect((await s.sheets())[0]!.autoFilter?.criteria).toEqual(criteria())
    const w = await mountSheet(s)
    expect(w.find('[data-xs-af="2"]').classes()).toContain('xs-on')
    await w.find('[data-xs-af="2"]').trigger('click')
    await settle()
    const choices = w.findAll('.xs-value').filter(l => l.find('.xs-value-text').text() !== 'Select all')
    expect(choices.map(l => [l.find('.xs-value-text').text(), (l.find('input').element as HTMLInputElement).checked])).toEqual([['a', true], ['b', false]])
    w.unmount()
  })

  it('Clear removes the saved criteria and reveals the rows, undo restores both', async () => {
    const s = source(book())
    const w = await mountSheet(s)
    ;(w.vm as unknown as SheetViewApi).clearFilter()
    await settle()
    expect((await s.book()).sheets[0]!.autoFilter?.criteria?.columns?.some(c => c.filter)).toBeFalsy()
    expect((await s.book()).sheets[0]!.rows?.['3']?.hidden).toBeFalsy()
    await s.undo()
    await settle()
    expect(w.find('[data-xs-af="2"]').classes()).toContain('xs-on')
    expect((await s.book()).sheets[0]!.rows?.['3']?.hidden).toBe(true)
    w.unmount()
  })

  it('writes a changed selection and its hidden rows to xlsx, and synchronizes another viewer', async () => {
    const s = source(book())
    const a = await mountSheet(s)
    const b = await mountSheet(s)
    await a.find('[data-xs-af="2"]').trigger('click')
    await settle()
    for (const label of a.findAll('.xs-value').filter(l => ['a', 'b'].includes(l.find('.xs-value-text').text()))) await label.find('input').trigger('change')
    await a.find('[data-xs-apply]').trigger('click')
    await until(() => b.find('[data-cell="3:1"]').exists())
    const opened = await openXlsx((await saveXlsx(s)).data)
    expect(opened.book.sheets[0]!.autoFilter?.criteria?.columns?.[0]?.filter).toEqual({ type: 'values', values: ['b'] })
    expect(opened.book.sheets[0]!.rows?.['2']?.hidden).toBe(true)
    expect(opened.book.sheets[0]!.rows?.['3']?.hidden).toBeFalsy()
    a.unmount(); b.unmount()
  })

  it('Reapply uses the saved supported conditions after an edit and keeps them in xlsx', async () => {
    const s = source(book())
    const w = await mountSheet(s)
    await s.edit([{ sheet: 'S', cell: 'B3', text: 'a' }])
    await settle()
    expect(await (w.vm as unknown as SheetViewApi).reapplyFilter()).toBe(true)
    await settle()
    expect((await s.book()).sheets[0]!.rows?.['3']?.hidden).toBeFalsy()
    const read = await openXlsx((await saveXlsx(s)).data)
    expect(read.book.sheets[0]!.autoFilter?.criteria).toEqual(criteria())
    w.unmount()
  })

  it('shows unsupported saved criteria without replacing them or attempting to reapply them', async () => {
    const saved: FilterCriteria = { columns: [{ column: 1, filter: { type: 'top', value: 10 } }] }
    const s = source(book(saved))
    const w = await mountSheet(s)
    expect(w.find('[data-xs-af="2"]').classes()).toContain('xs-on')
    await w.find('[data-xs-af="2"]').trigger('click')
    await settle()
    expect(w.find('[data-xs-stored-filter]').exists()).toBe(true)
    expect(w.find('[data-xs-apply]').attributes('disabled')).toBeDefined()
    expect(await (w.vm as unknown as SheetViewApi).reapplyFilter()).toBe(false)
    expect((await openXlsx((await saveXlsx(s)).data)).book.sheets[0]!.autoFilter?.criteria).toEqual(saved)
    w.unmount()
  })

  it('retains button settings and a complex sort when a supported condition changes', async () => {
    const saved: FilterCriteria = { columns: [{ column: 0, hiddenButton: true, filter: { type: 'values', values: ['one'] } }, ...criteria().columns!], sort: { range: 'A2:B3', conditions: [{ range: 'A2:A3', descending: true }, { range: 'B2:B3' }] } }
    const s = source(book(saved))
    const w = await mountSheet(s)
    expect(w.find('[data-xs-af="1"]').exists()).toBe(false)
    expect(w.find('[data-xs-filter-status]').text()).toContain('A, B')
    await s.setFilterCriteria('S', { ...saved, columns: [saved.columns![0]!, { column: 1, filter: { type: 'values', blank: true } }] })
    const reopened = await openXlsx((await saveXlsx(s)).data)
    expect(reopened.book.sheets[0]!.autoFilter?.criteria).toEqual({ ...saved, columns: [saved.columns![0]!, { column: 1, filter: { type: 'values', blank: true } }] })
    w.unmount()
  })

  it('applies an escaped literal substring and keeps its wildcard syntax after saving', async () => {
    const saved: FilterCriteria = { columns: [{ column: 1, filter: { type: 'custom', conditions: [{ operator: 'equal', value: '*a~*b*' }] } }] }
    const input = book(saved)
    input.sheets[0]!.cells.B2.value = 'a*b'
    const s = source(input)
    const w = await mountSheet(s)
    await w.find('[data-xs-af="2"]').trigger('click')
    await settle()
    expect((w.find('.xs-popover-input').element as HTMLInputElement).value).toBe('a*b')
    expect(await (w.vm as unknown as SheetViewApi).reapplyFilter()).toBe(true)
    expect((await openXlsx((await saveXlsx(s)).data)).book.sheets[0]!.autoFilter?.criteria).toEqual(saved)
    w.unmount()
  })
})
