<script setup lang="ts">
// The marks over the cells, in the scroller's content: the copied range's moving border, the fill handle at the selection's
// corner and the outline of a fill being dragged. A mark in frozen panes goes over them, as the editor does.
import { computed } from 'vue'
import { useSheetContext } from './context'

const ctx = useSheetContext()
const { clip, fill } = ctx
const m = ctx.messages

const place = (b: { top: number; left: number; width: number; height: number; frozen: boolean }) => ({
  top: `${b.top}px`,
  left: `${b.left}px`,
  width: `${b.width}px`,
  height: `${b.height}px`,
  zIndex: b.frozen ? '7' : '1',
})
const ants = computed(() => (clip.ants.value ? place(clip.ants.value) : null))
const frame = computed(() => (fill.frame.value ? place(fill.frame.value) : null))
/** The fill handle keeps its size at any zoom, so it can still be grabbed zoomed out: its own zoom undoes the sheet's, and its place is in the resulting pixels. */
const handle = computed(() => {
  const h = fill.handle.value
  const z = ctx.zoom.value
  if (!h) return null
  return { top: `${h.y * z}px`, left: `${h.x * z}px`, zIndex: h.frozen ? '7' : '1', ...(z !== 1 ? { zoom: String(1 / z) } : {}) }
})
</script>

<template>
  <div v-if="ants" class="xs-ants" :style="ants" aria-hidden="true" data-xs-ants></div>
  <div v-if="frame" class="xs-fill-frame" :style="frame" aria-hidden="true" data-xs-fill-frame></div>
  <div
    v-if="handle"
    :class="['xs-fill-handle', ctx.v.touch.value ? 'xs-fill-touch' : '', fill.handle.value?.dragging ? 'xs-fill-dragging' : '']"
    :style="handle"
    :title="m.fillHandle"
    data-xs-fill-handle
    @pointerdown="fill.onDown"
    @pointermove="fill.onMove"
    @pointerup="fill.onUp"
    @pointercancel="fill.onCancel"
    @lostpointercapture="fill.onCancel"
    @dblclick.stop="fill.autofill()"
  ></div>
</template>
