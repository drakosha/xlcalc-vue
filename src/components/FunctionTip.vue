<script setup lang="ts">
// Excel's argument tooltip below the editor (above it where the viewport has no room below), with the current argument in
// bold, kept inside the viewport and wrapped in a narrow one; a press keeps the editor's focus.
import { computed, onMounted, ref, watch } from 'vue'
import type { TipPart } from '../core/function-tip'
import { clampLeft, placeBelowOrAbove, type PopupRect } from './popup'

const props = withDefaults(
  defineProps<{
    /** The editor, in the coordinates of the element the tooltip is placed in. */
    rect: PopupRect
    parts: readonly TipPart[]
    /** The part in view, in the same coordinates: the tooltip stays inside it. */
    view?: PopupRect
  }>(),
  { view: undefined },
)

/** A narrow viewport wraps the tooltip after a separator or an opening parenthesis (ru-RU's `;` has no space to break at), not inside a name. */
const BREAK = /[;,(]$/
const pieces = (text: string) => text.split(/(?<=[;,(])/)
const el = ref<HTMLElement | null>(null)
/** The tooltip's measured size: it positions itself once it knows its text's width and height. */
const size = ref({ width: 0, height: 0 })
function measure(): void {
  const e = el.value
  if (e) size.value = { width: e.offsetWidth, height: e.offsetHeight }
}
onMounted(measure)
watch(() => [props.parts, props.view?.width], measure, { flush: 'post' })

const box = computed(() => {
  const r = props.rect
  const v = props.view
  const { width, height } = size.value
  return { top: `${placeBelowOrAbove(r, height, v)}px`, left: `${clampLeft(r.left, width, v)}px`, ...(v ? { maxWidth: `${v.width}px` } : {}) }
})
</script>

<template>
  <div ref="el" class="xs-ft" role="tooltip" :style="box" data-xs-tip @mousedown.prevent>
    <template v-for="(p, i) in parts" :key="i">
      <b v-if="p.current" class="xs-ft-on">{{ p.text }}</b>
      <template v-else>
        <template v-for="(piece, j) in pieces(p.text)" :key="j">{{ piece }}<wbr v-if="BREAK.test(piece)" /></template>
      </template>
    </template>
  </div>
</template>

<style src="../autocomplete.css"></style>
