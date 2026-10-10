// A row's or a column's style reaches past the data, as Excel draws it: the grid asks for what it draws, not only the data,
// and a source that cuts its windows at the data's edge does not make it ask again and again.
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { SheetSource, WindowRequest } from '../src/source'
import { cellAddress } from '../src/core/a1'
import { gridWindow, mountSheet, settle } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

/** A wide frame, as a desktop's: the grid draws columns far past three columns of data. */
const wide = () =>
  vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
    return this.classList.contains('xs-scroller') ? 1920 : 0
  })

/** Three columns of data and a header row with a fill of its own (a row style): every cell of row 1 has it, empty or not. */
function lined(cut: boolean) {
  const asked: WindowRequest[] = []
  const source: SheetSource = {
    sheets: async () => [{ name: 'S', maxRow: 10, maxCol: 3 }],
    window: async (req) => {
      asked.push(req)
      const cols = cut ? Math.max(0, Math.min(req.cols, 3 - req.col + 1)) : req.cols
      const w = gridWindow({ ...req, cols }, (r, c) => (r <= 10 && c <= 3 ? `${r}:${c}` : ''))
      w.styles = [{}, { fill: { color: '#d9e1f2' } }]
      for (let c = req.col; c < req.col + cols; c++) w.styleIds[cellAddress(1, c)] = 1
      return w
    },
    edit: async () => ({ ok: true, changes: [] }),
  }
  return { source, asked }
}
describe('styles of rows and columns past the data', () => {
  it('a header row’s fill runs across the whole frame, past the last column of data', async () => {
    wide()
    const { source } = lined(false)
    const w = await mountSheet(source)
    const fill = (c: number) => (w.find(`[data-cell="1:${c}"]`).element as HTMLElement).style.backgroundColor
    expect(fill(3)).toBe('#d9e1f2')
    expect(fill(20)).toBe('#d9e1f2')
    expect(fill(30)).toBe('#d9e1f2')
  })

  it('a source cutting its windows at the data does not make the grid ask again on every scroll', async () => {
    wide()
    const { source, asked } = lined(true)
    const w = await mountSheet(source)
    const scroller = w.find('.xs-scroller')
    const before = asked.length
    for (const top of [40, 80, 120, 160]) {
      scroller.element.scrollTop = top
      await scroller.trigger('scroll')
      await settle()
    }
    expect(asked.length).toBe(before)
    expect(w.find('[data-cell="2:3"]').text()).toBe('2:3')
  })
})

describe('the source’s edge and edits past it', () => {
  it('a value typed past the edge the source cut at comes back with the next window', async () => {
    wide()
    let maxCol = 3
    const asked: WindowRequest[] = []
    const typed = new Map<string, string>()
    const source: SheetSource = {
      sheets: async () => [{ name: 'S', maxRow: 10, maxCol }],
      window: async (req) => {
        asked.push(req)
        const cols = Math.max(0, Math.min(req.cols, maxCol - req.col + 1))
        return gridWindow({ ...req, cols }, (r, c) => typed.get(cellAddress(r, c)) ?? (c <= 3 ? `${r}:${c}` : ''))
      },
      edit: async (entries) => {
        for (const e of entries) typed.set(e.cell, e.text ?? '')
        maxCol = 8
        return { ok: true, changes: entries.map((e) => ({ sheet: 'S', cell: e.cell, text: e.text ?? '', value: e.text })) }
      },
    }
    const w = await mountSheet(source)
    const address = w.find('.xs-address')
    await address.setValue('H2')
    await address.trigger('keydown', { key: 'Enter' })
    await settle()
    await w.find('.xs-scroller').trigger('keydown', { key: 'x' })
    await w.find('.xs-editor').trigger('keydown', { key: 'Enter' })
    await settle()
    expect(w.find('[data-cell="2:8"]').text()).toBe('x')
  })
})
