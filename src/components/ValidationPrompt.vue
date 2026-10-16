<script setup lang="ts">
// Excel's input message of a rule of data validation while its cell is selected: a pale yellow box below the cell, from
// its middle to the right, its title in bold; left of that, or over the cell, where the view has no room.
import { computed, nextTick, ref, watch } from 'vue'
import type { ValidationRect } from '../core/validation'
import { placeBelowOrAbove } from './popup'

const props = withDefaults(
  defineProps<{
    /** The active cell (a merged area whole), in the coordinates of the element the component is placed in. */
    rect: ValidationRect
    prompt: { title?: string; text: string }
    /** What of the grid is in view, in the same coordinates. */
    view?: ValidationRect
  }>(),
  { view: undefined },
)

/** The gap between the cell and the box. */
const GAP = 4
const el = ref<HTMLElement | null>(null)
/** The box's size once drawn, for the room it needs. */
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
  // past the view's right: the box ends at the cell's middle, or at the view's edge when that is nearer
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
