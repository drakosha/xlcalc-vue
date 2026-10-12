<script setup lang="ts">
// The marks over the cells, in the scroller's content: the copied range's moving border, the fill handle at the
// selection's corner and the frame of a fill being dragged. A mark in frozen panes goes over them, as the field does.
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
const handle = computed(() => {
  const h = fill.handle.value
  return h ? { top: `${h.y}px`, left: `${h.x}px`, zIndex: h.frozen ? '7' : '1' } : null
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
