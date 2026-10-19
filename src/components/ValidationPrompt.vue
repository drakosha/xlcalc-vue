<script setup lang="ts">
// Excel's input message for a data validation rule while its cell is selected: a pale yellow box below the cell, from its
// middle to the right, with its title in bold; shifted left, or over the cell, where the viewport has no room.
import { computed, nextTick, ref, watch } from 'vue'
import { placeBelowOrAbove, type PopupRect } from './popup'

const props = withDefaults(
  defineProps<{
    /** The active cell (an entire merged range), in the coordinates of the element the component is placed in. */
    rect: PopupRect
    prompt: { title?: string; text: string }
    /** The part of the grid in view, in the same coordinates. */
    view?: PopupRect
  }>(),
  { view: undefined },
)

/** The gap between the cell and the box. */
const GAP = 4
const el = ref<HTMLElement | null>(null)
/** The box's size once drawn, for the space it needs. */
const size = ref({ width: 0, height: 0 })
watch(
  () => [props.prompt, props.rect, props.view],
  async () => {
    await nextTick()
    if (el.value) size.value = { width: el.value.offsetWidth, height: el.value.offsetHeight }
  },
  { immediate: true, flush: 'post' },
)

const place = computed(() => {
  const r = props.rect
  const v = props.view
  const { width, height } = size.value
  let left = r.left + r.width / 2
  // beyond the viewport's right edge: the box ends at the cell's middle, or at the viewport's edge when that is closer
  if (v && left + width > v.left + v.width) left = Math.max(v.left + GAP, Math.min(left, v.left + v.width - GAP) - width)
  const top = placeBelowOrAbove(r, height, v, GAP)
  return { left: `${left}px`, top: `${top}px`, ...(v ? { maxWidth: `${Math.max(120, v.width - 2 * GAP)}px` } : {}) }
})
</script>

<template>
  <div ref="el" class="xs-vp" role="tooltip" :style="place" data-xs-prompt>
    <b v-if="prompt.title">{{ prompt.title }}</b>
    <span v-if="prompt.text" class="xs-vp-text">{{ prompt.text }}</span>
  </div>
</template>

<style src="../validation.css"></style>
