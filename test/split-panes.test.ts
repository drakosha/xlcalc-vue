// Excel's split window drawn: two or four panes laid out as Excel draws them, each reading its own windows; panes in a row of
// panes scroll up and down together and in a column of panes left and right together (measured in Excel for Mac); a press
// makes its pane active, which takes the keys; the bars move and remove the lines; Freeze Panes freezes at the lines.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { SheetViewApi } from '../src'
import type { EditOutcome, SheetInfo, SheetSplit } from '../src/contract/source'
import { colPx } from '../src/core/geometry'
import { paneLayout, SPLIT_BAR_PX } from '../src/core/panes'
import { active, editorOpen, fakeSource, mountSheet, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
/** The default row (15 pt) and column on screen, and the headings: the numbers' 56 px and the letters' 24 px. */
const ROW = 20
const COL = colPx(null)
const LEAD = 56
const HEAD = 24

function book(split: SheetSplit | undefined, extra: Partial<SheetInfo> = {}) {
  const sheets: SheetInfo[] = [{ name: 'S', maxRow: 300, maxCol: 40, ...(split ? { split } : {}), ...extra }]
  const f = fakeSource(sheets)
  f.source.split = vi.fn(async (_sheet: string, panes: SheetSplit | null): Promise<EditOutcome> => {
    const { split: _, ...rest } = sheets[0]!
    sheets[0] = panes ? { ...rest, split: panes } : rest
    return { ok: true, changes: [] }
  })
  f.source.freeze = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [] }))
  f.source.structure = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [] }))
  return f
}
/** A split after `rows` rows and `cols` columns of the window from A1, in the file's points. */
const lines = (rows: number, cols: number) => ({ ...(cols ? { x: ((LEAD + cols * COL) * 3) / 4 } : {}), ...(rows ? { y: ((HEAD + rows * ROW) * 3) / 4 } : {}) })
const pane = (w: VueWrapper, id: number) => w.find(`[data-xs-pane="${id}"]`)
const el = (w: VueWrapper, id: number) => pane(w, id).element as HTMLElement
async function scroll(w: VueWrapper, id: number, to: { top?: number; left?: number }): Promise<void> {
  if (to.top !== undefined) el(w, id).scrollTop = to.top
  if (to.left !== undefined) el(w, id).scrollLeft = to.left
  await pane(w, id).trigger('scroll')
  await settle()
}

describe("the panes' layout", () => {
  const input = { width: 800, height: 600, headH: 24, leadW: 56, barV: 15, barH: 15 }

  it('places four panes after the lines and their bars; a bottom pane hides the letters and a right one the numbers, and only the right and bottom panes show scrollbars', () => {
    const { panes, x, y } = paneLayout({ ...input, x: 200, y: 150 })
    expect([x, y]).toEqual([200, 150])
    const after = 200 + SPLIT_BAR_PX
    const below = 150 + SPLIT_BAR_PX
    expect(panes.map((p) => [p.id, p.box])).toEqual([
      [0, { left: 0, top: 0, width: 200, height: 150 }],
      [1, { left: after, top: 0, width: 800 - after, height: 150 }],
      [2, { left: 0, top: below, width: 200, height: 600 - below }],
      [3, { left: after, top: below, width: 800 - after, height: 600 - below }],
    ])
    // the top left pane hides both scrollbars past its edges; the top right one its horizontal one, the bottom left one its vertical one
    expect(panes.map((p) => p.scroller)).toEqual([
      { left: 0, top: 0, width: 200 + 15, height: 150 + 15 },
      { left: -56, top: 0, width: 800 - after + 56, height: 150 + 15 },
      { left: 0, top: -24, width: 200 + 15, height: 600 - below + 24 },
      { left: -56, top: -24, width: 800 - after + 56, height: 600 - below + 24 },
    ])
    // panes that scroll together have the same client area: the top ones the same height, the left ones the same width
    const client = panes.map((p) => p.client)
    expect(client[0]!.height).toBe(client[1]!.height)
    expect(client[2]!.height).toBe(client[3]!.height)
    expect(client[0]!.width).toBe(client[2]!.width)
    expect(client[1]!.width).toBe(client[3]!.width)
    expect(client[0]).toEqual({ width: 200, height: 150 })
  })

  it('has a top and a bottom pane with a horizontal line alone, a left and a right one with a vertical line, and keeps a line inside the window', () => {
    const h = paneLayout({ ...input, x: null, y: 150 })
    expect(h.panes.map((p) => p.id)).toEqual([0, 2])
    // both show a vertical scrollbar, only the bottom one a horizontal one, as Excel draws them
    expect(h.panes.map((p) => [p.scroller.width, p.scroller.height])).toEqual([
      [800, 150 + 15],
      [800, 600 - 150 - SPLIT_BAR_PX + 24],
    ])
    const v = paneLayout({ ...input, x: 200, y: null })
    expect(v.panes.map((p) => p.id)).toEqual([0, 1])
    expect(v.panes.map((p) => [p.scroller.width, p.scroller.height])).toEqual([
      [200 + 15, 600],
      [800 - 200 - SPLIT_BAR_PX + 56, 600],
    ])
    expect(paneLayout({ ...input, x: 2000, y: null }).x).toBe(800 - SPLIT_BAR_PX)
  })
})

describe('a split window on the sheet', () => {
  it('draws four panes, the top left one from A1 and the bottom right one from the split’s top left cell, each read in windows of its own', async () => {
    const f = book({ ...lines(4, 2), topLeft: 'AD400', activePane: 'bottomRight' }, { maxRow: 1000, maxCol: 60 })
    Object.assign(f.source, { maxCells: 3000 })
    const w = await mountSheet(f.source)
    expect(w.findAll('.xs-pane')).toHaveLength(4)
    expect(w.findAll('[data-xs-split]').map((b) => b.attributes('data-xs-split'))).toEqual(['x', 'y', 'xy'])
    // the top left pane starts at A1, the bottom right one at AD400: its rows and columns from there are drawn
    expect(pane(w, 0).find('[data-cell="1:1"]').text()).toBe('S:A1')
    expect(pane(w, 3).find('[data-cell="400:30"]').text()).toBe('S:AD400')
    expect(pane(w, 3).find('[data-cell="1:1"]').exists()).toBe(false)
    // the bottom left pane: the bottom rows in the left columns; the top right one: the top rows in the right columns
    expect(pane(w, 2).find('[data-cell="400:1"]').text()).toBe('S:A400')
    expect(pane(w, 1).find('[data-cell="1:30"]').text()).toBe('S:AD1')
    const covers = (row: number, col: number) => f.requests.some((r) => row >= r.row && row < r.row + r.rows && col >= r.col && col < r.col + r.cols)
    expect([covers(1, 1), covers(400, 30), covers(400, 1), covers(1, 30)]).toEqual([true, true, true, true])
    // each request within the source's limit on cells, one for each pane
    expect(f.requests.every((r) => r.rows * r.cols <= 3000)).toBe(true)
    expect(f.requests).toHaveLength(4)
  })

  it('scrolls a row of panes up and down together and a column of panes left and right together, as Excel does', async () => {
    const w = await mountSheet(book({ ...lines(4, 2), topLeft: 'C5' }).source)
    await scroll(w, 0, { top: 400 })
    expect([el(w, 0).scrollTop, el(w, 1).scrollTop, el(w, 2).scrollTop, el(w, 3).scrollTop]).toEqual([400, 400, 80, 80])
    await scroll(w, 3, { left: 300 })
    expect([el(w, 0).scrollLeft, el(w, 1).scrollLeft, el(w, 2).scrollLeft, el(w, 3).scrollLeft]).toEqual([0, 300, 0, 300])
    await scroll(w, 2, { top: 1000, left: 50 })
    expect([el(w, 0).scrollTop, el(w, 1).scrollTop, el(w, 2).scrollTop, el(w, 3).scrollTop]).toEqual([400, 400, 1000, 1000])
    expect([el(w, 0).scrollLeft, el(w, 1).scrollLeft, el(w, 2).scrollLeft, el(w, 3).scrollLeft]).toEqual([50, 300, 50, 300])
    // the panes draw what their own scroll shows: the bottom ones row 51 at their top (1000 px of 20 px rows)
    expect(pane(w, 3).find('[data-cell="51:12"]').exists()).toBe(true)
    expect(pane(w, 1).find('[data-cell="51:12"]').exists()).toBe(false)
  })

  it('makes the pane pressed in active: it takes the keys and the editor, and the selection moves in it', async () => {
    const w = await mountSheet(book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' }).source)
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['3'])
    expect([0, 1, 2, 3].map((id) => pane(w, id).attributes('tabindex'))).toEqual(['-1', '-1', '-1', '0'])
    expect(pane(w, 3).find('.xs-editor').exists()).toBe(true)
    // a press on a pane's scrollbars (the scroller itself) only scrolls it
    await pane(w, 1).trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['3'])
    await pane(w, 0).find('[data-cell="2:1"]').trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    await settle()
    expect(active(w)).toBe('2:1')
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['0'])
    // the panes stay where they were
    expect([el(w, 0).scrollTop, el(w, 2).scrollTop, el(w, 1).scrollLeft]).toEqual([0, 4 * ROW, 2 * COL])
    expect(pane(w, 0).find('.xs-editor').exists()).toBe(true)
    expect(pane(w, 3).find('.xs-editor').exists()).toBe(false)
    // keys on another pane do nothing; on the active one they move the active cell
    await pane(w, 3).trigger('keydown', { key: 'ArrowDown' })
    expect(active(w)).toBe('2:1')
    await pane(w, 0).trigger('keydown', { key: 'ArrowDown' })
    await settle()
    expect(active(w)).toBe('3:1')
  })

  it('makes the pane active on a press that leaves the selection, such as on its fill handle', async () => {
    const w = await mountSheet(book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' }).source)
    await pane(w, 0).find('[data-xs-fill-handle]').trigger('pointerdown', { button: 0, pointerType: 'mouse', clientX: 0, clientY: 0 })
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['0'])
  })

  it('keeps an edit’s pane while a press in another points at a reference, and makes active the pane a press ends an edit in', async () => {
    const f = book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' })
    const w = await mountSheet(f.source)
    const field = () => w.find('.xs-editor').element as HTMLTextAreaElement
    await pane(w, 3).trigger('keydown', { key: '=' })
    await type(w, '=')
    field().setSelectionRange(1, 1)
    await pane(w, 0).find('[data-cell="2:1"]').trigger('pointerdown', { button: 0, pointerId: 1, pointerType: 'mouse' })
    await settle()
    expect(field().value).toBe('=A2')
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['3'])
    await type(w, 'abc')
    await pane(w, 0).find('[data-cell="3:1"]').trigger('pointerdown', { button: 0, pointerId: 2, pointerType: 'mouse' })
    await settle()
    expect(editorOpen(w)).toBe(false)
    expect(f.edits.at(-1)).toEqual([{ sheet: 'S', cell: 'A1', text: 'abc' }])
    expect(active(w)).toBe('3:1')
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['0'])
  })

  it('scrolls the active pane to the active cell, and the panes sharing its rows or columns with it', async () => {
    const w = await mountSheet(book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' }).source)
    await pane(w, 0).find('[data-cell="2:1"]').trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    await settle()
    for (let i = 0; i < 30; i++) await pane(w, 0).trigger('keydown', { key: 'ArrowDown' })
    await settle()
    expect(active(w)).toBe('32:1')
    // the top panes moved down to row 32, the bottom ones stayed
    expect(el(w, 0).scrollTop).toBeGreaterThan(0)
    expect(el(w, 1).scrollTop).toBe(el(w, 0).scrollTop)
    expect(el(w, 2).scrollTop).toBe(4 * ROW)
  })

  it('moves a line dragged by its bar to a row’s edge and keeps it through the source; a double click removes it, as does a drag to the window’s edge', async () => {
    const f = book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' })
    const w = await mountSheet(f.source)
    const bar = w.find('[data-xs-split="y"]')
    await bar.trigger('pointerdown', { button: 0, clientY: HEAD + 4 * ROW })
    // dropped 7 px past row 8's bottom edge: the line snaps to that edge
    await bar.trigger('pointermove', { clientY: HEAD + 8 * ROW + 7 })
    await bar.trigger('pointerup')
    await settle()
    expect(f.source.split).toHaveBeenLastCalledWith('S', { ...lines(8, 2), topLeft: 'C5', activePane: 'bottomRight' })
    await w.find('[data-xs-split="x"]').trigger('dblclick')
    await settle()
    expect(f.source.split).toHaveBeenLastCalledWith('S', { ...lines(8, 0), topLeft: 'A5', activePane: 'bottomLeft' })
    expect(w.findAll('.xs-pane')).toHaveLength(2)
    const h = w.find('[data-xs-split="y"]')
    await h.trigger('pointerdown', { button: 0, clientY: HEAD + 8 * ROW })
    await h.trigger('pointermove', { clientY: 10 })
    await h.trigger('pointerup')
    await settle()
    expect(f.source.split).toHaveBeenLastCalledWith('S', null)
    expect(w.find('.xs-split').exists()).toBe(false)
    expect(w.findAll('.xs-scroller')).toHaveLength(1)
  })

  it('removed, leaves the window at the top left pane’s scroll, as Excel does', async () => {
    const w = await mountSheet(book({ ...lines(4, 2), topLeft: 'C5', activePane: 'bottomRight' }).source)
    await scroll(w, 0, { top: 400, left: 120 })
    expect(api(w).split(false)).toBe(false)
    await settle()
    const single = w.find('.xs-scroller').element as HTMLElement
    expect([single.scrollTop, single.scrollLeft]).toEqual([400, 120])
  })

  it('freezes at the lines with Freeze Panes, as Excel does, whatever the active cell', async () => {
    const f = book({ ...lines(4, 2), topLeft: 'C5' })
    const w = await mountSheet(f.source)
    await pane(w, 0).find('[data-cell="2:1"]').trigger('pointerdown', { button: 0, pointerType: 'mouse' })
    await settle()
    await api(w).freezePanes()
    expect(f.source.freeze).toHaveBeenLastCalledWith('S', { rows: 4, cols: 2 })
  })

  it('keeps the panes that are not active on their top row when the heights of rows above them arrive', async () => {
    const f = book({ y: lines(4, 0).y, topLeft: 'A200' }, { maxRow: 400 })
    // rows 1 to 99 are 30 points (40 px) high
    const read = f.source.window.bind(f.source)
    f.source.window = async (req, signal) => {
      const got = await read(req, signal)
      if ('pending' in got) return got
      const rowHeights: Record<number, number> = {}
      for (let r = req.row; r < Math.min(100, req.row + req.rows); r++) rowHeights[r] = 30
      return { ...got, rowHeights }
    }
    const w = await mountSheet(f.source)
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['0'])
    expect(el(w, 2).scrollTop).toBe(99 * 40 + 100 * ROW)
    expect(pane(w, 2).find('[data-cell="200:1"]').exists()).toBe(true)
  })

  it('goes from a split sheet to one without a split and back, the keys going to the active pane', async () => {
    const sheets: SheetInfo[] = [
      { name: 'S', maxRow: 300, maxCol: 40, split: { ...lines(4, 0), topLeft: 'A50', activePane: 'bottomLeft' } },
      { name: 'T', maxRow: 300, maxCol: 40 },
    ]
    const w = await mountSheet(fakeSource(sheets).source)
    await w.findAll('.xs-tab').find((t) => t.text() === 'T')!.trigger('click')
    await settle()
    expect(w.find('.xs-split').exists()).toBe(false)
    await w.find('.xs-scroller').trigger('keydown', { key: 'ArrowDown' })
    expect(active(w)).toBe('2:1')
    await w.findAll('.xs-tab').find((t) => t.text() === 'S')!.trigger('click')
    await settle()
    expect(w.findAll('.xs-pane').map((p) => p.find('.xs-scroller').attributes('data-xs-pane'))).toEqual(['0', '2'])
    expect(w.findAll('.xs-pane-active').map((p) => p.attributes('data-xs-pane'))).toEqual(['2'])
    expect([el(w, 0).scrollTop, el(w, 2).scrollTop]).toEqual([0, 49 * ROW])
    await pane(w, 2).trigger('keydown', { key: 'ArrowRight' })
    expect(active(w)).toBe('1:2')
  })

  it('is drawn as one pane without a split, and with frozen panes, which take its place', async () => {
    const w = await mountSheet(book(undefined).source)
    expect(w.find('.xs-split').exists()).toBe(false)
    expect(w.findAll('.xs-scroller')).toHaveLength(1)
    const frozen = await mountSheet(book({ ...lines(4, 2), topLeft: 'C5' }, { freeze: { rows: 1, cols: 0 } }).source)
    expect(frozen.find('.xs-split').exists()).toBe(false)
  })
})
