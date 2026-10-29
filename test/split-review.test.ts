// The split window's edges: a renamed sheet keeps its scroll; windows no pane draws any more stop serving cells; a reference
// or the fill handle dragged into another pane takes the cell under the pointer there; the keys follow the active pane when a
// line goes.
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { EditOutcome, SheetInfo, SheetOp } from '../src/contract/source'
import { fakeSource, goTo, mountSheet, settle, type } from './helpers'
import { api, book, COL, el, HEAD, LEAD, lines, pane, ROW, scroll } from './split-fixture'

afterEach(() => {
  document.body.innerHTML = ''
})

/** Waits out a repaint's delay. */
async function repainted(): Promise<void> {
  await settle(10)
  await new Promise((r) => setTimeout(r, 150))
  await settle(10)
}

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

describe('windows no pane draws', () => {
  /** A source whose A1:B2 is merged while `merged` holds. */
  function merging() {
    const state = { merged: true }
    const f = book(undefined)
    const read = f.source.window.bind(f.source)
    f.source.window = async (req, signal) => {
      const got = await read(req, signal)
      if ('pending' in got) return got
      return { ...got, merges: state.merged && req.row <= 1 && req.col <= 1 ? ['A1:B2'] : [] }
    }
    return { f, state }
  }

  it('give way to a split made in the session: a repainted cell is drawn anew in its pane', async () => {
    let ver = 1
    const f = book(undefined, {}, { cell: (sh, r, c) => `${sh}:${r}:${c}:v${ver}` })
    const w = await mountSheet(f.source)
    await w.find('[data-cell="1:3"]').trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    await settle()
    expect(api(w).split()).toBe(true)
    await settle()
    expect(w.findAll('.xs-pane')).toHaveLength(2)
    expect(pane(w, 0).find('[data-cell="1:1"]').text()).toBe('S:1:1:v1')
    ver = 2
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1:B2' }] })
    await repainted()
    expect(pane(w, 0).find('[data-cell="1:1"]').text()).toBe('S:1:1:v2')
  })

  it('go with a split removed: an unmerge repainted is drawn', async () => {
    const { f, state } = merging()
    const w = await mountSheet(f.source)
    await w.find('[data-cell="3:3"]').trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    await settle()
    expect(api(w).split()).toBe(true)
    await settle()
    expect(w.findAll('.xs-pane')).toHaveLength(4)
    expect(api(w).split(false)).toBe(false)
    await settle()
    state.merged = false
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1:B2' }] })
    await repainted()
    expect(w.find('[data-cell="1:2"]').exists()).toBe(true)
  })

  it('go with frozen panes unfrozen: an unmerge repainted is drawn', async () => {
    const { f, state } = merging()
    const sheets: SheetInfo[] = [{ name: 'S', maxRow: 300, maxCol: 40 }]
    f.source.sheets = async () => sheets
    f.source.freeze = vi.fn(async (_s: string, panes: { rows: number; cols: number } | null): Promise<EditOutcome> => {
      sheets[0] = panes ? { name: 'S', maxRow: 300, maxCol: 40, freeze: panes } : { name: 'S', maxRow: 300, maxCol: 40 }
      f.emit({ kind: 'reload' })
      return { ok: true, changes: [] }
    })
    const w = await mountSheet(f.source)
    await w.find('[data-cell="3:3"]').trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    await settle()
    await api(w).freezePanes()
    await settle(10)
    await api(w).freezePanes('unfreeze')
    await settle(10)
    state.merged = false
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1:B2' }] })
    await repainted()
    expect(w.find('[data-cell="1:2"]').exists()).toBe(true)
  })
})

describe('a drag into another pane', () => {
  /** Every pane at the page's top left, as happy-dom lays them out: a point is in the pane its element is in. */
  const at = (row: number, col: number) => ({ clientX: LEAD + (col - 0.5) * COL, clientY: HEAD + (row - 0.5) * ROW })

  it('of a reference ends at the cell under the pointer in that pane', async () => {
    const w = await mountSheet(book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' }).source)
    const field = () => w.find('.xs-editor').element as HTMLTextAreaElement
    await pane(w, 3).trigger('keydown', { key: '=' })
    await type(w, '=')
    field().setSelectionRange(1, 1)
    await pane(w, 0).find('[data-cell="2:1"]').trigger('pointerdown', { button: 0, pointerId: 1, pointerType: 'mouse', ...at(2, 1) })
    await settle()
    expect(field().value).toBe('=A2')
    await pane(w, 0).find('[data-cell="4:1"]').trigger('pointermove', { pointerId: 1, pointerType: 'mouse', buttons: 1, ...at(4, 1) })
    await settle()
    expect(field().value).toBe('=A2:A4')
  })

  it('of the fill handle fills up to the cell under the pointer in that pane', async () => {
    const w = await mountSheet(book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' }).source)
    await pane(w, 3).find('[data-cell="5:3"]').trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    await settle()
    const target = pane(w, 2).find('[data-cell="5:1"]').element
    // the handle holds the pointer, so the cell under it is found by its place on the page
    vi.spyOn(document, 'elementFromPoint').mockReturnValue(target)
    const handle = pane(w, 3).find('[data-xs-fill-handle]')
    await handle.trigger('pointerdown', { button: 0, pointerId: 3, pointerType: 'mouse' })
    await handle.trigger('pointermove', { pointerId: 3, pointerType: 'mouse', ...at(5, 1) })
    await settle()
    // the outline from A5 to C5: its left edge at column A
    expect(pane(w, 3).find('[data-xs-fill-frame]').attributes('style')).toContain(`left: ${LEAD}px`)
    vi.restoreAllMocks()
  })
})

describe('the keys when a line goes', () => {
  async function focused() {
    const f = book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' })
    const w = await mountSheet(f.source)
    await pane(w, 3).find('[data-cell="6:4"]').trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    await settle()
    ;(pane(w, 3).element as HTMLElement).focus()
    expect(document.activeElement).toBe(pane(w, 3).element)
    return { f, w }
  }

  it('go to the pane that becomes active when the bar is double-clicked', async () => {
    const { w } = await focused()
    await w.find('[data-xs-split="x"]').trigger('dblclick')
    await settle()
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['2'])
    expect(document.activeElement).toBe(pane(w, 2).element)
  })

  it('go to the pane that becomes active when the bar is dragged off', async () => {
    const { w } = await focused()
    const bar = w.find('[data-xs-split="x"]')
    await bar.trigger('pointerdown', { button: 0, clientX: LEAD + 2 * COL })
    await bar.trigger('pointermove', { clientX: 10 })
    await bar.trigger('pointerup')
    await settle()
    expect(document.activeElement).toBe(pane(w, 2).element)
  })

  it('go to the single window when the split is removed or gives way to frozen panes', async () => {
    const one = await focused()
    api(one.w).split(false)
    await settle()
    expect(document.activeElement).toBe(one.w.find('.xs-scroller').element)
    document.body.innerHTML = ''
    const two = await focused()
    const sheets: SheetInfo[] = [{ name: 'S', maxRow: 300, maxCol: 40, freeze: { rows: 4, cols: 2 } }]
    two.f.source.freeze = vi.fn(async (): Promise<EditOutcome> => {
      two.f.source.sheets = async () => sheets
      return { ok: true, changes: [] }
    })
    await api(two.w).freezePanes()
    await settle()
    expect(two.w.find('.xs-split').exists()).toBe(false)
    expect(document.activeElement).toBe(two.w.find('.xs-scroller').element)
  })
})
