// Excel's Clear Formats and Clear All as the sheet's methods: one call to the source's `clear` over the selection, then the windows over it read again.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook, type SheetViewApi } from '../src'
import type { WindowRequest } from '../src/contract/source'
import { cellText, goTo, mountSheet, settle } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({
  styles: [{ font: { bold: true }, fill: { color: '#FFF2CC' } }],
  sheets: [
    {
      name: 'S',
      cells: { A1: 'Name', B1: 'Status', A2: { value: 'n2', style: 0, note: { text: 'a note' } }, B2: 'a', A3: { value: 'n3', style: 0 }, B3: 'b', A4: { value: 'n4', style: 0 }, B4: 'a' },
    },
  ],
})
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi

describe('Clear Formats and Clear All', () => {
  it('Clear Formats sends the selection to `clear` and reads its windows again: the cells keep their values and lose their style', async () => {
    const source = memorySource(book())
    const clear = vi.spyOn(source, 'clear')
    const window = vi.spyOn(source, 'window')
    const w = await mountSheet(source)
    await goTo(w, 'A2:A3')
    const before = window.mock.calls.length
    const out = await api(w).clear('formats')
    await settle()
    expect(out?.ok).toBe(true)
    expect(clear).toHaveBeenCalledWith('S', ['A2:A3'], 'formats')
    expect(window.mock.calls.length).toBeGreaterThan(before)
    expect(cellText(w, 2, 1)).toBe('n2')
    expect(w.find('[data-cell="2:1"]').attributes('style') ?? '').not.toContain('font-weight')
    expect(w.find('[data-cell="4:1"]').attributes('style')).toContain('font-weight')
  })

  it('Clear All removes the contents and the note', async () => {
    const source = memorySource(book())
    const w = await mountSheet(source)
    await goTo(w, 'A2:B2')
    await api(w).clear('all')
    await settle()
    expect(cellText(w, 2, 1)).toBe('')
    expect(cellText(w, 2, 2)).toBe('')
    expect(w.find('[data-cell="2:1"]').attributes('title') ?? '').not.toContain('a note')
  })

  it('an entire column is sent as Excel writes it', async () => {
    const source = memorySource(book())
    const clear = vi.spyOn(source, 'clear')
    const w = await mountSheet(source)
    await w.findAll('.xs-head')[0]!.trigger('click')
    await api(w).clear('formats')
    expect(clear).toHaveBeenCalledWith('S', ['A:A'], 'formats')
  })

  it('under a display filter only the runs of visible rows are cleared', async () => {
    const source = memorySource(book())
    const clear = vi.spyOn(source, 'clear')
    const w = await mountSheet(source)
    await w.findAll('.xs-head')[1]!.find('.xs-funnel').trigger('click')
    await settle()
    const label = w.findAll('.xs-value').find((l) => l.find('.xs-value-text').text() === 'b')!
    await label.find('input').trigger('change')
    await w.find('[data-xs-apply]').trigger('click')
    await settle()
    await goTo(w, 'A1:A4')
    await api(w).clear('formats')
    expect(clear).toHaveBeenCalledWith('S', ['A1:A2', 'A4'], 'formats')
  })

  it('sends nothing on a read-only sheet, without `clear`, or over a locked cell', async () => {
    const ro = await mountSheet(memorySource(book()), { readOnly: true })
    expect(await api(ro).clear('all')).toBeNull()
    const bare = memorySource(book())
    delete (bare as { clear?: unknown }).clear
    expect(await api(await mountSheet(bare)).clear('all')).toBeNull()
    const locked = memorySource({ sheets: [{ ...book().sheets[0]!, locked: [{ range: 'A2', reason: 'Managed by a script' }] }] })
    const clear = vi.spyOn(locked, 'clear')
    const w = await mountSheet(locked)
    await goTo(w, 'A1:A3')
    expect(await api(w).clear('all')).toBeNull()
    await settle()
    expect(clear).not.toHaveBeenCalled()
    expect(w.find('.xs-message').text()).toContain('Managed by a script')
  })

  it('reads its windows again even when the source sends no repaint', async () => {
    const source = memorySource(book())
    const w = await mountSheet(source)
    const window = vi.spyOn(source, 'window')
    await goTo(w, 'A2')
    await api(w).clear('formats')
    await settle()
    const asked = window.mock.calls.map((c) => c[0] as WindowRequest)
    expect(asked.some((r) => r.sheet === 'S' && r.row <= 2 && r.row + r.rows > 2)).toBe(true)
  })
})
