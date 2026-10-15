<script setup lang="ts">
// Excel's tip with a function's arguments under the field being edited (over it where the view has no room below), the
// current argument in bold, kept inside the view and wrapped on a narrow one; a press keeps the field's focus.
import { computed, onMounted, ref, watch } from 'vue'
import type { TipPart } from '../core/function-tip'
import type { ValidationRect } from '../core/validation'

const props = withDefaults(
  defineProps<{
    /** The field, in the coordinates of the element the tip is placed in. */
    rect: ValidationRect
    parts: readonly TipPart[]
    /** What is in view, in the same coordinates: the tip stays inside it. */
    view?: ValidationRect
  }>(),
  { view: undefined },
)

/** A narrow view wraps the tip after a separator or an opening bracket (ru-RU's `;` has no space to break at), not inside a name. */
const BREAK = /[;,(]$/
const pieces = (text: string) => text.split(/(?<=[;,(])/)
const el = ref<HTMLElement | null>(null)
/** The tip's measured size: it places itself once it knows how wide and high its text is. */
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
  const below = r.top + r.height
  const up = !!v && below + height > v.top + v.height && r.top - height >= v.top
  const left = v ? Math.max(v.left, Math.min(r.left, v.left + v.width - width)) : r.left
  return { top: `${up ? r.top - height : below}px`, left: `${left}px`, ...(v ? { maxWidth: `${v.width}px` } : {}) }
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
