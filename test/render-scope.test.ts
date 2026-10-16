// What a change redraws: typing, events and moves touch only the cells they change, not every drawn cell.
import type { VueWrapper } from '@vue/test-utils'
import { describe, expect, it } from 'vitest'
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
})
