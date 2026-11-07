<script setup lang="ts">
// The AutoFilter's buttons, Excel's arrows at the bottom right of its header cells over its columns: a click opens the column's
// list; a column with a condition shows a funnel, a sorted one its direction. In the scroller's content, over frozen panes where the row is frozen.
import { computed } from 'vue'
import { useSheetContext } from './context'
import '../autofilter.css'

const ctx = useSheetContext()
const { v, geo, filter } = ctx
const m = ctx.messages

/** Excel's button is about 16 px square at 100%; 24 px on a touch screen, within the cell either way. */
const sizeOf = computed(() => (v.touch.value ? 24 : 16))

/** The buttons drawn: the header row's cells in view within the AutoFilter's columns. */
const buttons = computed(() => {
  const r = filter.auto.value
  if (!r || !filter.available.value) return []
  const d = ctx.indexOf(r.top)
  if (d === undefined || geo.heightOf(d) <= 0 || (d > v.frozenRows.value && (d < v.firstRow.value || d > v.lastRow.value))) return []
  const shown = [...v.frozenColList.value, ...v.shownCols.value].filter((c) => c >= r.left && c <= r.right && filter.canFunnel(c))
  return shown.map((c) => {
    const b = v.boxOf(d, c, d, c)
    const size = Math.min(sizeOf.value, Math.max(8, b.height - 2), Math.max(8, b.width - 2))
    const style = { top: `${b.top + b.height - size - 1}px`, left: `${b.left + b.width - size - 1}px`, width: `${size}px`, height: `${size}px`, zIndex: b.frozen ? '7' : '1' }
    const sort = filter.sortAt(c)
    return { c, style, on: filter.hasColumn(c), sort }
  })
})
function open(e: MouseEvent, c: number): void {
  void filter.openFunnel(c, (e.currentTarget as HTMLElement).getBoundingClientRect())
}
</script>

<template>
  <button
    v-for="b in buttons"
    :key="b.c"
    type="button"
    :class="['xs-af', { 'xs-on': b.on || b.sort, 'xs-open': filter.open.value === b.c }]"
    :style="b.style"
    :title="b.on ? m.filterOn : m.filter"
    :aria-label="b.on ? m.filterOn : m.filter"
    aria-haspopup="dialog"
    :aria-expanded="filter.open.value === b.c"
    :data-xs-af="b.c"
    data-xs-funnel
    tabindex="-1"
    @pointerdown.stop
    @mousedown.prevent
    @click.stop="open($event, b.c)"
  >
    <svg viewBox="0 0 12 12" aria-hidden="true">
      <path v-if="b.on" d="M1.5 2h9L7.2 6v4.2L4.8 9V6z" fill="currentColor" />
      <template v-else-if="b.sort">
        <path :d="b.sort === 'asc' ? 'M3 10.5V2.5M1.3 4.2 3 2.5l1.7 1.7' : 'M3 1.5v8M1.3 7.8 3 9.5l1.7-1.7'" fill="none" stroke="currentColor" stroke-width="1.1" />
        <path d="M6.5 5h4.5L8.75 8z" fill="currentColor" />
      </template>
      <path v-else d="M3 4.5h6L6 8z" fill="currentColor" />
    </svg>
  </button>
</template>
