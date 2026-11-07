import { afterEach, describe, expect, it } from 'vitest'
import { closeAll, source } from './local-fixture'
import { mountSheet, settle } from './helpers'

afterEach(async () => { await closeAll(); document.body.innerHTML = '' })
const runs = [{ start: 0, font: { bold: true, underline: 'double' as const } }, { start: 2, font: { italic: true, color: '#FF0000' } }]

describe('formatted runs on the sheet', () => {
  it('uses UTF-16 offsets and resets a cell font for each whole run font', async () => {
    const s = source({ styles: [{ font: { bold: true, underline: 'single' } }], sheets: [{ name: 'S', cells: { A1: { value: '😀<tag>', richText: runs, style: 0 } } }] })
    const w = await mountSheet(s)
    const spans = w.findAll('[data-cell="1:1"] .xs-rich-run')
    expect(spans.map(x => x.text())).toEqual(['😀', '<tag>'])
    expect((spans[0]!.element as HTMLElement).style.fontWeight).toBe('700')
    expect((spans[0]!.element as HTMLElement).style.textDecorationStyle).toBe('double')
    expect((spans[1]!.element as HTMLElement).style.fontWeight).toBe('400')
    expect((spans[1]!.element as HTMLElement).style.fontStyle).toBe('italic')
    expect((spans[1]!.element as HTMLElement).style.textDecorationLine).toBe('none')
    expect(w.find('[data-cell="1:1"] tag').exists()).toBe(false)
    w.unmount()
  })

  it('removes stale run spans on edit and restores them on undo without reopening the sheet', async () => {
    const s = source({ sheets: [{ name: 'S', cells: { A1: { value: '😀text', richText: runs } } }] })
    const w = await mountSheet(s)
    expect(w.findAll('.xs-rich-run')).toHaveLength(2)
    await s.edit([{ sheet: 'S', cell: 'A1', text: 'plain' }])
    await settle()
    expect(w.find('[data-cell="1:1"]').text()).toBe('plain')
    expect(w.find('.xs-rich-run').exists()).toBe(false)
    await s.undo()
    await settle()
    expect(w.findAll('.xs-rich-run').map(x => x.text())).toEqual(['😀', 'text'])
    w.unmount()
  })

  it.each(['link', 'merge', 'rotation', 'fill'] as const)('keeps the runs in %s content', async kind => {
    const s = source({ styles: [{ alignment: { ...(kind === 'rotation' ? { rotation: 45 } : {}), ...(kind === 'fill' ? { horizontal: 'fill' } : {}) } }], sheets: [{ name: 'S', cells: { A1: { value: '😀text', richText: runs, style: 0, ...(kind === 'link' ? { link: { url: 'https://example.org' } } : {}) } }, ...(kind === 'merge' ? { merges: ['A1:B1'] } : {}) }] })
    const w = await mountSheet(s)
    expect(w.findAll('[data-cell="1:1"] .xs-rich-run').length).toBeGreaterThanOrEqual(2)
    expect(w.find('[data-cell="1:1"]').text()).toContain('😀text')
    w.unmount()
  })
})
