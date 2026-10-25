<script setup lang="ts">
// Excel's zoom on the status bar: Zoom Out, the slider (10% to 100% on its left half, 100% to 400% on its right), Zoom In,
// and the percentage, which opens the Zoom dialog. Dragging the slider zooms as it moves and reports the zoom on release.
import { computed } from 'vue'
import { sliderOfZoom, stepZoom, zoomOfSlider } from '../core/zoom'
import { useSheetContext } from './context'
import { IconMinus, IconPlus } from './icons'

const ctx = useSheetContext()
const m = ctx.messages
const zm = ctx.zoomer
const pos = computed(() => sliderOfZoom(zm.percent.value))
const at = (e: Event) => zoomOfSlider(Number((e.target as HTMLInputElement).value))
</script>

<template>
  <span class="xs-zoom-bar" data-xs-zoom-bar>
    <button type="button" class="xs-tool" :title="m.zoomOut" :aria-label="m.zoomOut" data-xs-zoom-out @click="zm.choose(stepZoom(zm.percent.value, -1))">
      <IconMinus />
    </button>
    <input
      class="xs-zoom-slider"
      type="range"
      min="0"
      max="100"
      step="any"
      :value="pos"
      :aria-label="m.zoomSlider"
      :aria-valuetext="`${zm.percent.value}%`"
      :title="m.zoomSlider"
      data-xs-zoom-slider
      @input="zm.apply(at($event))"
      @change="zm.choose(at($event))"
    />
    <button type="button" class="xs-tool" :title="m.zoomIn" :aria-label="m.zoomIn" data-xs-zoom-in @click="zm.choose(stepZoom(zm.percent.value, 1))">
      <IconPlus />
    </button>
    <button type="button" class="xs-zoom-level" :title="m.zoomLevel" data-xs-zoom-level @click="zm.dialogOpen.value = true">{{ zm.percent.value }}%</button>
  </span>
</template>

<style src="../view.css"></style>
