// What a change redraws: typing, events and moves touch only the cells they change, not every drawn cell.
import type { VueWrapper } from '@vue/test-utils'
import { describe, expect, it, vi } from 'vitest'
import { nextTick } from 'vue'
import { sheetContextKey } from '../src/components/context'
import type { SheetContext } from '../src/core/sheet'
import { cellText, fakeSource, mountSheet, settle, type } from './helpers'

const contextOf = (w: VueWrapper) => (w.vm.$ as unknown as { provides: Record<symbol, SheetContext> }).provides[sheetContextKey as symbol]!

/** Counts the grid's calls of `cellView` (the template reads it from the context on each render). */
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
  it('typing in the cell field redraws no cell, a formula only when its references change', async () => {
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

  it('a change of a cell off screen redraws no cell, nor one that changes nothing; it shows once scrolled to', async () => {
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
    expect(count.calls).toBeGreaterThan(0)
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

  it("a stream of changes asks the active cell's rule once it stops", async () => {
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
})
