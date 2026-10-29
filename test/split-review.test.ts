// The split window's edges: a renamed sheet keeps its scroll.
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { EditOutcome, SheetInfo, SheetOp } from '../src/contract/source'
import { fakeSource, goTo, mountSheet, settle } from './helpers'
import { api, el, lines, ROW, scroll } from './split-fixture'

afterEach(() => {
  document.body.innerHTML = ''
})

describe('a renamed sheet', () => {
  function renamable(extra: Partial<SheetInfo> = {}) {
    const sheets: SheetInfo[] = [{ name: 'A', maxRow: 1000, maxCol: 40, ...extra }, { name: 'B', maxRow: 10, maxCol: 10 }]
    const f = fakeSource(sheets)
    f.source.sheets = async () => sheets.map((x) => ({ ...x }))
    f.source.structure = vi.fn(async (op: SheetOp): Promise<EditOutcome> => {
      if (op.op === 'renameSheet') {
        const i = sheets.findIndex((x) => x.name === op.sheet)
        sheets[i] = { ...sheets[i]!, name: op.name }
      }
      f.emit({ kind: 'reload' })
      return { ok: true, changes: [] }
    })
    return f
  }

  it('keeps the scroll the grid draws by', async () => {
    const w = await mountSheet(renamable().source)
    await goTo(w, 'A500')
    const sc = w.find('.xs-scroller').element as HTMLElement
    const top = sc.scrollTop
    expect(top).toBeGreaterThan(400 * ROW)
    await api(w).structure({ op: 'renameSheet', sheet: 'A', name: 'Data' })
    await settle()
    expect(sc.scrollTop).toBe(top)
    expect(w.find('[data-cell="500:1"]').exists()).toBe(true)
  })

  it('keeps the panes’ scroll', async () => {
    const w = await mountSheet(renamable({ split: { ...lines(4, 0), topLeft: 'A300', activePane: 'bottomLeft' } }).source)
    expect(el(w, 2).scrollTop).toBe(299 * ROW)
    await scroll(w, 2, { top: 400 * ROW })
    await scroll(w, 0, { top: 50 * ROW })
    await api(w).structure({ op: 'renameSheet', sheet: 'A', name: 'Data' })
    await settle()
    expect(w.findAll('.xs-tab').map((t) => t.text())).toEqual(['Data', 'B'])
    expect([el(w, 0).scrollTop, el(w, 2).scrollTop]).toEqual([50 * ROW, 400 * ROW])
  })
})
