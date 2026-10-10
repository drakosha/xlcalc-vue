// Windows: one request per slot in flight, a late answer for a previous sheet is dropped, frozen panes get their own slots,
// changes go into the loaded windows only, and the source's events patch or reload.
import { afterEach, describe, expect, it } from 'vitest'
import { cellText, fakeSource, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

describe('windows', () => {
  it('sends no second request for a slot while one is in flight', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 5000, maxCol: 10 }], { manual: true })
    const w = await mountSheet(f.source, { readOnly: true })
    expect(f.requests).toHaveLength(1)
    expect(f.requests[0]).toMatchObject({ sheet: 'S', row: 1, col: 1 })
    const scroller = w.find('.xs-scroller')
    for (const top of [2000, 4000, 6000, 8000]) {
      scroller.element.scrollTop = top
      await scroller.trigger('scroll')
      await settle(2)
    }
    expect(f.requests).toHaveLength(1)
    f.release(f.held[0]!)
    await settle()
    // The answer came: the scroll went far past it, so one new request for where the frame is now.
    expect(f.requests).toHaveLength(2)
    expect(f.requests[1]!.row).toBeGreaterThan(200)
    expect(f.requests[1]!.row).toBeLessThanOrEqual(400)
  })

  it('drops a late answer for the previous sheet', async () => {
    const f = fakeSource(
      [
        { name: 'One', maxRow: 50, maxCol: 5 },
        { name: 'Two', maxRow: 50, maxCol: 5 },
      ],
      { manual: true },
    )
    const w = await mountSheet(f.source, { readOnly: true })
    const first = f.held[0]!
    expect(first.req.sheet).toBe('One')
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    const second = f.held.find((p) => p.req.sheet === 'Two')!
    f.release(second)
    await settle()
    expect(cellText(w, 1, 1)).toBe('Two:A1')
    // The first sheet's answer arrives after: it must not overwrite the second sheet.
    f.release(first)
    await settle()
    expect(cellText(w, 1, 1)).toBe('Two:A1')
    expect(cellText(w, 2, 3)).toBe('Two:C2')
  })

  it('loads frozen rows and columns in their own slots', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 3000, maxCol: 40, freeze: { rows: 2, cols: 1 } }])
    const w = await mountSheet(f.source, { readOnly: true })
    const scroller = w.find('.xs-scroller')
    scroller.element.scrollTop = 20_000
    scroller.element.scrollLeft = 1_000
    await scroller.trigger('scroll')
    await settle()
    const kinds = f.requests.slice(1)
    expect(kinds.some((q) => q.row === 1 && q.rows === 2)).toBe(true)
    expect(kinds.some((q) => q.col === 1 && q.cols === 1 && q.row > 500)).toBe(true)
    expect(cellText(w, 1, 1)).toBe('S:A1')
    expect(cellText(w, 2, 20)).toBe('S:T2')
  })

  it('applies changes to the loaded windows only', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 5000, maxCol: 5 }], {
      edit: async () => ({
        ok: true,
        changes: [
          { sheet: 'S', cell: 'A1', text: 'mine', value: 'mine' },
          { sheet: 'S', cell: 'B2', text: '84', value: 84, formula: 'A2*2' },
          { sheet: 'S', cell: 'B4900', text: 'far away', value: 'far away' },
        ],
      }),
    })
    const w = await mountSheet(f.source)
    await w.find('.xs-scroller').trigger('keydown', { key: 'x' })
    await w.find('.xs-editor').trigger('keydown', { key: 'Enter' })
    await settle()
    expect(cellText(w, 1, 1)).toBe('mine')
    expect(cellText(w, 2, 2)).toBe('84')
    expect(w.find('[data-cell="2:2"]').attributes('title')).toContain('=A2*2')
    // The far cell was not loaded: its window comes from the source when it is needed.
    const scroller = w.find('.xs-scroller')
    scroller.element.scrollTop = 4890 * 20
    await scroller.trigger('scroll')
    await settle()
    expect(cellText(w, 4900, 2)).toBe('S:B4900')
  })

  it('takes changed events into the loaded windows, and reloads on a reload event', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 100, maxCol: 5 }])
    const w = await mountSheet(f.source, { readOnly: true })
    f.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: 'C3', text: 'from elsewhere', value: 'from elsewhere' }, { sheet: 'Other', cell: 'C4', text: 'x', value: 'x' }] })
    await settle()
    expect(cellText(w, 3, 3)).toBe('from elsewhere')
    expect(cellText(w, 4, 3)).toBe('S:C4')
    const before = f.requests.length
    f.emit({ kind: 'reload', message: 'Updated elsewhere' })
    await settle()
    expect(f.sheetsCalls).toBe(2)
    expect(f.requests.length).toBeGreaterThan(before)
    expect(cellText(w, 3, 3)).toBe('S:C3')
    expect(w.find('.xs-message').text()).toContain('Updated elsewhere')
  })

  it('shows the fallback slot when the book cannot be read, and retries', async () => {
    let fail = true
    const f = fakeSource([{ name: 'S', maxRow: 5, maxCol: 5 }])
    const sheets = f.source.sheets
    f.source.sheets = async () => {
      if (fail) throw new Error('no book')
      return sheets()
    }
    const w = await mountSheet(f.source)
    expect(w.find('.xs-fallback').text()).toContain('no book')
    fail = false
    await w.find('.xs-fallback button').trigger('click')
    await settle()
    expect(w.find('.xs-fallback').exists()).toBe(false)
    expect(cellText(w, 1, 1)).toBe('S:A1')
  })
})
