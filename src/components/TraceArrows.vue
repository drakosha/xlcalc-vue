<script setup lang="ts">
// Excel's tracer arrows over the cells, in the scroller's content so they scroll with the grid: a dot in the traced cell, a
// line and an arrowhead at the dependent, a dashed line to a worksheet icon for dependents on other sheets. Arrows between
// body cells go under the frozen panes and the headers; an arrow with an end in frozen panes goes over them, clipped to the cells' area.
import { computed, useId } from 'vue'
import type { DrawnArrow } from '../core/trace'
import { useSheetContext } from './context'

const ctx = useSheetContext()
const { trace, v } = ctx
const clipId = `${useId()}-trace-clip`

/** The arrowhead's length and half width, in px. */
const HEAD_LEN = 9
const HEAD_HALF = 4

const head = (a: DrawnArrow): string => {
  const dx = a.x2 - a.x1
  const dy = a.y2 - a.y1
  const len = Math.hypot(dx, dy) || 1
  const ux = dx / len
  const uy = dy / len
  const bx = a.x2 - ux * HEAD_LEN
  const by = a.y2 - uy * HEAD_LEN
  return `${a.x2},${a.y2} ${bx - uy * HEAD_HALF},${by + ux * HEAD_HALF} ${bx + uy * HEAD_HALF},${by - ux * HEAD_HALF}`
}
/** The two layers: under the frozen panes, and over them for arrows with an end in them. */
const layers = computed(() => [
  { id: 'body', z: 1, arrows: trace.drawn.value.filter((a) => !a.frozen) },
  { id: 'frozen', z: 7, arrows: trace.drawn.value.filter((a) => a.frozen) },
])
</script>

<template>
  <template v-for="layer in layers" :key="layer.id">
    <svg v-if="layer.arrows.length" class="xs-trace" :style="{ zIndex: layer.z }" width="1" height="1" aria-hidden="true" :data-xs-trace="layer.id">
      <defs v-if="layer.id === 'frozen'">
        <clipPath :id="clipId" clipPathUnits="userSpaceOnUse">
          <rect :x="v.viewBox.value.left" :y="v.viewBox.value.top" :width="v.viewBox.value.width" :height="v.viewBox.value.height" />
        </clipPath>
      </defs>
      <g :clip-path="layer.id === 'frozen' ? `url(#${clipId})` : undefined">
        <g v-for="a in layer.arrows" :key="a.key" :class="`xs-trace-${a.kind}`" :data-xs-arrow="a.key">
          <line :x1="a.x1" :y1="a.y1" :x2="a.x2" :y2="a.y2" />
          <circle :cx="a.x1" :cy="a.y1" r="2.5" />
          <polygon v-if="a.kind !== 'sheet'" :points="head(a)" />
          <g v-else :transform="`translate(${a.x2 - 7} ${a.y2 - 6})`" class="xs-trace-icon">
            <rect x="0.5" y="0.5" width="14" height="12" />
            <path d="M0.5 4.5h14M0.5 8.5h14M5.5 0.5v12M10 0.5v12" />
          </g>
        </g>
      </g>
    </svg>
  </template>
</template>

<style src="../trace.css"></style>
