<script setup lang="ts">
// Excel's split window (View > Split): its panes, each the grid scrolled apart, and the split bars between them. A bar is
// dragged to another row's or column's edge, and removed by a double click or by dragging it to the window's edge, as in Excel.
import { computed, onBeforeUnmount, onMounted, ref } from 'vue'
import { PT, SPLIT_BAR_PX, type Box } from '../core/panes'
import { useSheetContext } from './context'
import SheetGrid from './SheetGrid.vue'
import '../split.css'

const ctx = useSheetContext()
const { panes } = ctx
const el = ref<HTMLElement | null>(null)
const layout = computed(() => panes.layout.value)
const place = (b: Box) => ({ left: `${b.left}px`, top: `${b.top}px`, width: `${b.width}px`, height: `${b.height}px` })
const bar = `${SPLIT_BAR_PX}px`

// ── the window's size and the browser's scrollbars, which the panes' layout needs ──
let observer: ResizeObserver | null = null
/** The window's size, and the scrollbars once it is shown: a window mounted hidden measures none until then. */
function measure(): void {
  const host = el.value
  if (!host) return
  panes.box.value = { width: host.clientWidth, height: host.clientHeight }
  if (host.clientWidth > 0) {
    const bars = scrollbars(host)
    if (bars.v !== panes.scrollbars.value.v || bars.h !== panes.scrollbars.value.h) panes.scrollbars.value = bars
  }
}
/** The scrollbars' thickness: a scroller's size less its client area, 0 where they overlay the content. */
function scrollbars(host: HTMLElement): { v: number; h: number } {
  const probe = document.createElement('div')
  probe.className = 'xs-pane-scroller'
  probe.style.cssText = 'position:absolute;top:0;left:0;width:100px;height:100px;overflow:scroll;visibility:hidden'
  host.appendChild(probe)
  const out = { v: probe.offsetWidth - probe.clientWidth, h: probe.offsetHeight - probe.clientHeight }
  probe.remove()
  return out
}
onMounted(() => {
  if (!el.value) return
  measure()
  if (typeof ResizeObserver !== 'undefined') {
    observer = new ResizeObserver(measure)
    observer.observe(el.value)
  }
})
onBeforeUnmount(() => observer?.disconnect())

// ── the bars: a drag moves the lines live, snapped to an edge of a row or column, and the release keeps them ──
/** The lines a drag moves, the pointer's last place in the window, and whether it moved. */
let dragging: { x: boolean; y: boolean; at: { x: number; y: number }; moved: boolean } | null = null
/** The split's own lines in screen pixels, for the one a drag leaves where it is. */
const own = () => ({ x: panes.split.value?.x ? panes.split.value.x / PT : null, y: panes.split.value?.y ? panes.split.value.y / PT : null })
function down(e: PointerEvent, x: boolean, y: boolean): void {
  const lay = layout.value
  if (e.button !== 0 || !lay || !el.value) return
  ;(e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId)
  e.preventDefault()
  dragging = { x, y, at: { x: lay.x ?? 0, y: lay.y ?? 0 }, moved: false }
  panes.drag.value = { x: lay.x, y: lay.y }
}
function move(e: PointerEvent): void {
  const d = panes.drag.value
  if (!dragging || !d || !el.value) return
  const r = el.value.getBoundingClientRect()
  dragging.at = { x: e.clientX - r.left, y: e.clientY - r.top }
  dragging.moved = true
  // at the window's edge the bar follows the pointer, where the release removes it
  const to = (axis: 'x' | 'y') => (panes.offEdge(axis, dragging!.at[axis]) ? Math.max(0, dragging!.at[axis]) : panes.snap(axis, dragging!.at[axis]))
  panes.drag.value = { x: dragging.x ? to('x') : d.x, y: dragging.y ? to('y') : d.y }
}
function up(): void {
  const d = panes.drag.value
  const g = dragging
  dragging = null
  if (!g?.moved || !d) return void (panes.drag.value = null)
  const keep = own()
  const lines = { x: g.x ? (panes.offEdge('x', g.at.x) ? null : d.x) : keep.x, y: g.y ? (panes.offEdge('y', g.at.y) ? null : d.y) : keep.y }
  if (lines.x !== keep.x || lines.y !== keep.y) ctx.split.move(lines)
  panes.drag.value = null
}
function cancel(): void {
  dragging = null
  panes.drag.value = null
}
/** A double click removes the bar's line, both at their crossing. */
function remove(x: boolean, y: boolean): void {
  const keep = own()
  ctx.split.move({ x: x ? null : keep.x, y: y ? null : keep.y })
}
</script>

<template>
  <div ref="el" class="xs-split">
    <template v-if="layout">
      <div v-for="p in layout.panes" :key="p.id" class="xs-pane" :style="place(p.box)">
        <SheetGrid :pane="p.id" :style="place(p.scroller)" />
      </div>
      <div
        v-if="layout.x !== null"
        class="xs-split-bar xs-split-v"
        role="separator"
        aria-orientation="vertical"
        :style="{ left: `${layout.x}px`, width: bar }"
        data-xs-split="x"
        @pointerdown="down($event, true, false)"
        @pointermove="move"
        @pointerup="up"
        @pointercancel="cancel"
        @dblclick="remove(true, false)"
      ></div>
      <div
        v-if="layout.y !== null"
        class="xs-split-bar xs-split-h"
        role="separator"
        aria-orientation="horizontal"
        :style="{ top: `${layout.y}px`, height: bar }"
        data-xs-split="y"
        @pointerdown="down($event, false, true)"
        @pointermove="move"
        @pointerup="up"
        @pointercancel="cancel"
        @dblclick="remove(false, true)"
      ></div>
      <div
        v-if="layout.x !== null && layout.y !== null"
        class="xs-split-bar xs-split-cross"
        :style="{ left: `${layout.x}px`, top: `${layout.y}px`, width: bar, height: bar }"
        data-xs-split="xy"
        @pointerdown="down($event, true, true)"
        @pointermove="move"
        @pointerup="up"
        @pointercancel="cancel"
        @dblclick="remove(true, true)"
      ></div>
    </template>
  </div>
</template>
