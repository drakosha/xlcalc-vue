// What a change redraws: typing, events and moves affect only the cells they change, not every drawn cell.
import type { VueWrapper } from '@vue/test-utils'
import { describe, expect, it, vi } from 'vitest'
import { nextTick, triggerRef } from 'vue'
import { sheetContextKey } from '../src/components/context'
import type { SheetContext } from '../src/core/sheet'
import { cellText, fakeSource, mountSheet, settle, type } from './helpers'

const contextOf = (w: VueWrapper) => (w.vm.$ as unknown as { provides: Record<symbol, SheetContext> }).provides[sheetContextKey as symbol]!

/** Counts the grid's calls to `cellView` (the template reads it from the context on each render). */
function countCells(ctx: SheetContext) {
  const own = ctx.views.cellView
  let n = 0
  ctx.views.cellView = (...a) => (n++, own(...a))
  return { take: () => ((n = 0), undefined), get calls() {
    return n
  } }
}

async function open() {
  const f = fakeSource([{ name: 'S', maxRow: 200, maxCol: 20 }])
  const w = await mountSheet(f.source)
  const ctx = contextOf(w)
  return { w, ctx, f, count: countCells(ctx) }
}

describe('render scope', () => {
  it('typing in the cell editor redraws no cell, and a formula only when its references change', async () => {
    const { w, ctx, count } = await open()
    ctx.ed.startEdit({ mode: 'edit', text: '' })
    await settle()
    for (const text of ['h', 'he', 'hel']) {
      count.take()
      await type(w, text)
      await nextTick()
      expect(count.calls, text).toBe(0)
    }
    await type(w, '=A1+')
    await settle()
    for (const text of ['=A1+1', '=A1+12', '=A1+12*']) {
      count.take()
      await type(w, text)
      await nextTick()
      expect(count.calls, text).toBe(0)
    }
    count.take()
    await type(w, '=A1+12*B2')
    await nextTick()
    expect(count.calls).toBeGreaterThan(0)
    expect(w.find('[data-cell="2:2"]').attributes('style')).toContain('inset')
  })

  it('typing in the formula bar shows the draft in its cell and redraws no cell', async () => {
    const { w, ctx, count } = await open()
    ctx.ed.startEdit({ mode: 'edit', text: 'a', place: 'bar' })
    await settle()
    expect(cellText(w, 1, 1)).toBe('a')
    for (const text of ['ab', 'abc']) {
      count.take()
      ctx.editor.value!.text = text
      await nextTick()
      expect(count.calls, text).toBe(0)
      expect(cellText(w, 1, 1)).toBe(text)
    }
  })

  it('a change of a cell off screen redraws no cell, nor does one that changes nothing; it shows once scrolled into view', async () => {
    const { w, ctx, f, count } = await open()
    const d = ctx.v.lastRow.value + 20
    expect(ctx.w.slotAt(d, 3)).not.toBeNull()
    count.take()
    f.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: `C${d}`, text: 'far', value: 'far' }] })
    await nextTick()
    expect(count.calls).toBe(0)
    count.take()
    f.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: 'C2', text: 'S:C2', value: 'S:C2' }] })
    await nextTick()
    expect(count.calls).toBe(0)
    count.take()
    f.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: 'C2', text: 'near', value: 'near' }] })
    await nextTick()
    // Its row and the one above (whose bottom line reads it) are redrawn, not the grid.
    const perRow = ctx.v.shownCols.value.length
    expect(count.calls).toBeGreaterThan(0)
    expect(count.calls).toBeLessThanOrEqual(2 * perRow)
    expect(cellText(w, 2, 3)).toBe('near')
    ctx.scrollTop.value = (d - 5) * 20
    await settle()
    expect(cellText(w, d, 3)).toBe('far')
  })

  it('a change of the active cell off screen reaches the formula bar and the state', async () => {
    const { w, ctx, f } = await open()
    const d = ctx.v.lastRow.value + 20
    ctx.selected.value = { row: d, col: 3 }
    await settle()
    f.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: `C${d}`, text: 'far', value: 'far', formula: 'A1&"far"' }] })
    await settle()
    expect((w.find('.xs-fx-input').element as HTMLTextAreaElement).value).toBe('=A1&"far"')
    const states = w.emitted('state')!
    expect(states[states.length - 1]![0]).toMatchObject({ value: 'far', formula: 'A1&"far"' })
  })

  it("a burst of changes requests the active cell's rule once it ends", async () => {
    const f = fakeSource([{ name: 'S', maxRow: 200, maxCol: 20 }])
    const validation = vi.fn(async () => null)
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] })
    try {
      const w = await mountSheet({ ...f.source, validation })
      vi.advanceTimersByTime(100)
      await settle()
      validation.mockClear()
      for (let i = 0; i < 20; i++) f.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: `C${10 + i}`, text: 'x', value: 'x' }] })
      await settle()
      expect(validation).not.toHaveBeenCalled()
      vi.advanceTimersByTime(100)
      await settle()
      expect(validation).toHaveBeenCalledTimes(1)
      w.unmount()
    } finally {
      vi.useRealTimers()
    }
  })

  it('scrolling row by row redraws once per step of rows, with the screen always drawn with spare rows around it', async () => {
    const { ctx, count } = await open()
    const v = ctx.v
    let redraws = 0
    for (let i = 1; i <= 40; i++) {
      count.take()
      ctx.scrollTop.value = i * 20
      await nextTick()
      if (count.calls) redraws++
      const top = ctx.geo.rowAt(ctx.scrollTop.value)
      const bottom = ctx.geo.rowAt(ctx.scrollTop.value + ctx.viewH.value)
      expect(v.firstRow.value).toBeLessThanOrEqual(Math.max(1, top - 8))
      expect(v.lastRow.value).toBeGreaterThanOrEqual(Math.min(ctx.rowCount.value, bottom + 8))
    }
    expect(redraws).toBeLessThanOrEqual(11)
  })

  it('rows above the viewport that grow taller keep the viewport on its rows: the scroll moves by their growth', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 2000, maxCol: 10 }])
    let tall = false
    const window = f.source.window
    f.source.window = async (req, signal) => {
      const got = await window(req, signal)
      if (tall && 'text' in got) for (let r = req.row; r < req.row + got.text.length; r++) got.rowHeights[r] = 30
      return got
    }
    const w = await mountSheet(f.source, { readOnly: true })
    const ctx = contextOf(w)
    const el = w.find('.xs-scroller').element as HTMLElement
    el.scrollTop = 4000
    await w.find('.xs-scroller').trigger('scroll')
    await settle()
    const anchor = ctx.geo.rowAt(4000)
    const win = ctx.w.slots.body.value.win!
    tall = true
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1:J2000' }] })
    await settle()
    const grown = (anchor - win.row) * 20
    expect(ctx.scrollTop.value).toBe(4000 + grown)
    expect(el.scrollTop).toBe(4000 + grown)
    expect(ctx.geo.rowAt(ctx.scrollTop.value)).toBe(anchor)
  })

  it('a move of the active cell or of the selection redraws the rows it affects', async () => {
    const { ctx, count } = await open()
    const perRow = ctx.v.shownCols.value.length
    ctx.selected.value = { row: 5, col: 3 }
    await nextTick()
    count.take()
    ctx.selected.value = { row: 6, col: 3 }
    await nextTick()
    expect(count.calls).toBe(2 * perRow)
    count.take()
    ctx.extent.value = { row: 9, col: 5 }
    await nextTick()
    // Rows 6 to 9 enter the selection.
    expect(count.calls).toBe(4 * perRow)
    count.take()
    ctx.extent.value = { row: 10, col: 5 }
    await nextTick()
    // The old bottom row loses its edge, and the new one gets it.
    expect(count.calls).toBe(2 * perRow)
  })

  it('a scroll by a step draws only the rows that come into view', async () => {
    const { ctx, count } = await open()
    ctx.scrollTop.value = 200
    await settle()
    const before = ctx.v.shownRows.value
    count.take()
    ctx.scrollTop.value = 200 + 4 * 20
    await nextTick()
    const after = ctx.v.shownRows.value
    const come = after.filter((d) => !before.includes(d)).length
    expect(come).toBe(4)
    expect(count.calls).toBe(come * ctx.v.shownCols.value.length)
  })

  it("dragging a column's width or a row's height redraws no cell, or only the row's", async () => {
    const { ctx, count } = await open()
    const drawn = ctx.v.shownCols.value
    count.take()
    ctx.userWidths.value[3] = 72
    triggerRef(ctx.userWidths)
    await nextTick()
    // The same columns drawn, one a pixel wider.
    expect(ctx.v.shownCols.value).toEqual(drawn)
    expect(count.calls).toBe(0)
    count.take()
    ctx.geo.setHeights([[5, 33]])
    await nextTick()
    expect(count.calls).toBe(ctx.v.shownCols.value.length)
  })

  it('a horizontal scroll creates the views of the columns it brings into view, not of those that stay (numbers: no text overflows)', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 200, maxCol: 20 }], { cell: (_s, r, c) => String(r * 100 + c) })
    const ctx = contextOf(await mountSheet(f.source))
    const count = countCells(ctx)
    const before = new Set(ctx.v.shownCols.value)
    const rows = ctx.v.frozenRowList.value.length + ctx.v.shownRows.value.length
    count.take()
    ctx.scrollLeft.value = ctx.geo.offsets.value[4]!
    await nextTick()
    const added = ctx.v.shownCols.value.filter((c) => !before.has(c)).length
    expect(added).toBeGreaterThan(0)
    expect(count.calls).toBe(added * rows)
  })
})
