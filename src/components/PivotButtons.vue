<script setup lang="ts">
// The filter buttons of the displayed sheet's PivotTables, at the bottom right of Row Labels, Column Labels, a field's header
// or a Filters area's item as the source tells them: a click opens the field's filter menu; a field filtered shows a funnel.
// In the scroller's content, as the AutoFilter's buttons are.
import { computed } from 'vue'
import { parseCell } from 'xlcalc/a1'
import type { SheetPivotButton } from '../contract/source'
import { axisFields, filtered, type PivotTarget } from '../pivot-model'
import { useSheetContext } from './context'
import '../autofilter.css'

const ctx = useSheetContext()
const { v, geo } = ctx
const m = ctx.messages
const f = ctx.pivotFilter

const sizeOf = computed(() => (v.touch.value ? 24 : 16))
/** Whether the button's field, or in compact form any field of its axis, has a filter. */
const on = (t: PivotTarget, b: SheetPivotButton) => (b.axis === 'pages' || t.cells.buttons.filter((y) => y.axis === b.axis).length > 1 ? [b.field] : axisFields(t, b.axis)).some((x) => filtered(t, x))
const buttons = computed(() => {
  if (!ctx.pivotDialogs.can.value) return []
  const out: { key: string; t: PivotTarget; b: SheetPivotButton; style: Record<string, string>; on: boolean }[] = []
  for (const t of ctx.pivotDialogs.views.value) {
    for (const b of t.cells.buttons) {
      const at = parseCell(b.cell)
      const d = at ? ctx.indexOf(at.row) : undefined
      if (!at || d === undefined || geo.heightOf(d) <= 0 || geo.widthOf(at.col) <= 0) continue
      if (d > v.frozenRows.value && (d < v.firstRow.value || d > v.lastRow.value)) continue
      const box = v.boxOf(d, at.col, d, at.col)
      const size = Math.min(sizeOf.value, Math.max(8, box.height - 2), Math.max(8, box.width - 2))
      const style = { top: `${box.top + box.height - size - 1}px`, left: `${box.left + box.width - size - 1}px`, width: `${size}px`, height: `${size}px`, zIndex: box.frozen ? '7' : '1' }
      out.push({ key: `${t.pivot}:${b.cell}`, t, b, style, on: on(t, b) })
    }
  }
  return out
})
function open(e: MouseEvent, t: PivotTarget, b: SheetPivotButton): void {
  const r = (e.currentTarget as HTMLElement).getBoundingClientRect()
  f.openMenu(t, b.axis, b.field, { left: r.left, top: r.bottom })
}
</script>

<template>
  <button
    v-for="x in buttons"
    :key="x.key"
    type="button"
    :class="['xs-af', 'xs-pivot-button', { 'xs-on': x.on }]"
    :style="x.style"
    :title="x.on ? m.filterOn : m.filter"
    :aria-label="x.on ? m.filterOn : m.filter"
    aria-haspopup="dialog"
    :data-pivot-button="x.b.cell"
    tabindex="-1"
    @pointerdown.stop
    @mousedown.prevent
    @click.stop="open($event, x.t, x.b)"
  >
    <svg viewBox="0 0 12 12" aria-hidden="true">
      <path v-if="x.on" d="M1.5 2h9L7.2 6v4.2L4.8 9V6z" fill="currentColor" />
      <path v-else d="M3 4.5h6L6 8z" fill="currentColor" />
    </svg>
  </button>
</template>
