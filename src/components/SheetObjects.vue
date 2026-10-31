<script setup lang="ts">
// The sheet's slicers and timelines over the cells, in the scroller's content so they scroll and zoom with the grid; each
// pane of frozen panes draws its part of a box, and each pane of a split window its own. A press on the frame selects one
// and drags it, a handle resizes it, as Excel's do; Delete removes it, Esc gives the keys back to the grid, Alt+S and
// Alt+C are a slicer's Multi-Select and Clear Filter. The presses stay here: the grid under them does not see them.
import { computed, nextTick, ref, watch } from 'vue'
import type { SheetObjectRef } from '../contract/view'
import { boxOfPlace, partsOf, placeOfBox, type DrawnBox } from '../core/drawing'
import { useSheetContext } from './context'
import SlicerBox from './SlicerBox.vue'
import TimelineBox from './TimelineBox.vue'
import '../slicers.css'

const ctx = useSheetContext()
const { v, geo, slicers } = ctx

/** The smallest box a resize leaves, in px. */
const MIN_PX = 40
const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
type Handle = (typeof HANDLES)[number]

/** A move or a resize under way: the box it started from and how far the pointer went, in unzoomed px. */
const drag = ref<{ key: string; o: SheetObjectRef; handle: Handle | null; from: DrawnBox; x: number; y: number; dx: number; dy: number; moved: boolean } | null>(null)

/** The box a drag leaves: moved by the pointer, or its edges of the handle moved, no smaller than `MIN_PX`. */
function dragged(d: NonNullable<typeof drag.value>): DrawnBox {
  const b = { ...d.from }
  if (!d.handle) return { ...b, x: Math.max(0, b.x + d.dx), y: Math.max(0, b.y + d.dy) }
  if (d.handle.includes('e')) b.width = Math.max(MIN_PX, b.width + d.dx)
  if (d.handle.includes('s')) b.height = Math.max(MIN_PX, b.height + d.dy)
  if (d.handle.includes('w')) {
    const w = Math.max(MIN_PX, b.width - d.dx)
    b.x += b.width - w
    b.width = w
  }
  if (d.handle.includes('n')) {
    const h = Math.max(MIN_PX, b.height - d.dy)
    b.y += b.height - h
    b.height = h
  }
  return b
}

const rows = { rowOf: ctx.rowOf, indexOf: ctx.indexOf }
/** Every slicer and timeline with a place, as its pane draws its parts. */
const drawn = computed(() => {
  const list = [...slicers.slicers.value.map((x) => ({ o: { kind: 'slicer' as const, name: x.name }, place: x.place, slicer: x, timeline: null })), ...slicers.timelines.value.map((x) => ({ o: { kind: 'timeline' as const, name: x.name }, place: x.place, slicer: null, timeline: x }))]
  return list.flatMap((x) => {
    const at = x.place && boxOfPlace(x.place, geo, rows)
    if (!at) return []
    const key = `${x.o.kind}:${x.o.name}`
    const d = drag.value
    const box = d && d.key === key ? dragged(d) : at
    return partsOf(box, v).map((p) => ({ ...x, key, box, part: p }))
  })
})
const isSelected = (o: SheetObjectRef) => slicers.selected.value?.kind === o.kind && slicers.selected.value.name === o.name

// ── selecting, moving and resizing ──
/** A press on the frame (not on an item, a button or the periods) selects the object and starts a move; on a handle, a resize. */
function onDown(e: PointerEvent, x: (typeof drawn.value)[number], handle: Handle | null = null): void {
  if (e.button !== 0 && e.button !== 2) return
  slicers.select(x.o)
  void nextTick(() => (e.currentTarget as HTMLElement | null)?.closest<HTMLElement>('[data-xs-object]')?.focus({ preventScroll: true }))
  if (e.button !== 0 || (!handle && (e.target as HTMLElement).closest('button, [data-xs-periods], input'))) return
  if (!slicers.canEdit(x.slicer ?? x.timeline)) return
  ;(e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId)
  drag.value = { key: x.key, o: x.o, handle, from: x.box, x: e.clientX, y: e.clientY, dx: 0, dy: 0, moved: false }
}
function onMove(e: PointerEvent): void {
  const d = drag.value
  if (!d) return
  const z = ctx.zoom.value
  d.dx = (e.clientX - d.x) / z
  d.dy = (e.clientY - d.y) / z
  d.moved ||= Math.abs(d.dx) + Math.abs(d.dy) > 2
}
/** The release: a box moved or resized goes to its new place, as one step of undo. */
function onUp(): void {
  const d = drag.value
  if (!d) return
  if (!d.moved) return void (drag.value = null)
  const place = placeOfBox(dragged(d), geo, rows)
  void slicers.place(d.o, place).finally(() => (drag.value = null))
}
function onContext(e: MouseEvent, o: SheetObjectRef): void {
  e.preventDefault()
  slicers.select(o)
  ctx.menu.openObject(e.clientX, e.clientY)
}
/** The keys of a selected object: Delete removes it, Esc returns to the grid, Alt+S and Alt+C as Excel's header buttons. */
function onKey(e: KeyboardEvent, o: SheetObjectRef): void {
  const alt = e.altKey && !e.ctrlKey && !e.metaKey
  if (e.key === 'Delete' || e.key === 'Backspace') void slicers.remove(o)
  else if (e.key === 'Escape') {
    slicers.select(null)
    ctx.focusGrid()
  } else if (alt && e.code === 'KeyS' && o.kind === 'slicer') slicers.toggleMulti(o.name)
  else if (alt && e.code === 'KeyC') void slicers.clear(o)
  else return
  e.preventDefault()
}
// a drag of a box that was removed meanwhile ends
watch(drawn, (list) => {
  if (drag.value && !list.some((x) => x.key === drag.value!.key)) drag.value = null
})
</script>

<template>
  <div
    v-for="x in drawn"
    :key="`${x.key}:${x.part.key}`"
    :class="['xs-object', `xs-object-${x.o.kind}`, { 'xs-object-on': isSelected(x.o), 'xs-object-drag': drag?.key === x.key }]"
    :style="{ top: `${x.part.top}px`, left: `${x.part.left}px`, width: `${x.box.width}px`, height: `${x.box.height}px`, zIndex: x.part.z, clipPath: x.part.clip || undefined }"
    :tabindex="isSelected(x.o) ? 0 : -1"
    role="group"
    :aria-label="(x.slicer ?? x.timeline)!.caption ?? x.o.name"
    :data-xs-object="x.key"
    :data-xs-part="x.part.key || undefined"
    @pointerdown.stop="onDown($event, x)"
    @pointermove.stop="onMove"
    @pointerup.stop="onUp"
    @pointercancel.stop="drag = null"
    @mousedown.stop
    @click.stop
    @dblclick.stop
    @contextmenu.stop="onContext($event, x.o)"
    @keydown.stop="onKey($event, x.o)"
  >
    <SlicerBox v-if="x.slicer" :slicer="x.slicer" />
    <TimelineBox v-else-if="x.timeline" :timeline="x.timeline" />
    <template v-if="isSelected(x.o) && slicers.canEdit(x.slicer ?? x.timeline)">
      <span v-for="h in HANDLES" :key="h" :class="['xs-object-handle', `xs-object-handle-${h}`]" :data-xs-handle="h" @pointerdown.stop="onDown($event, x, h)" @pointermove.stop="onMove" @pointerup.stop="onUp"></span>
    </template>
  </div>
</template>
