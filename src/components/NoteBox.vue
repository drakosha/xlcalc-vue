<script setup lang="ts">
// The note being edited, anchored to its cell as Excel's note box: to the right of the cell (to its left where the viewport has
// no room), the author's name in bold on the first line, then the text. Esc or a press outside the box saves it, as Excel does.
// With `peek`, the note of the cell under the mouse in the same box, read-only, as Excel shows a note on hover.
import { computed, onMounted, ref } from 'vue'
import { useSheetContext } from './context'
import { useOutsidePress } from './popup'

const props = defineProps<{ peek?: boolean }>()
const ctx = useSheetContext()
const { v, views, notes } = ctx
const m = ctx.messages

/** The gap between the cell and the box, and the box's default width (as in the CSS). */
const GAP = 8
const WIDTH = 180
const box = ref<HTMLElement | null>(null)
const input = ref<HTMLTextAreaElement | null>(null)

/** The note shown: the one being edited, or with `peek` the hovered cell's. */
const shown = computed(() => (props.peek ? notes.peek.value : notes.draft.value))
const place = computed(() => {
  const d = shown.value
  const at = d && d.sheet === ctx.sheetName.value ? ctx.indexOf(d.row) : undefined
  if (!d || at === undefined) return null
  const sp = views.spanOf(d.row, d.col)
  const cell = v.boxOf(at, d.col, at + (sp?.rs ?? 1) - 1, d.col + (sp?.cs ?? 1) - 1)
  const view = v.viewBox.value
  const right = cell.left + cell.width + GAP
  const left = right + WIDTH > view.left + view.width && cell.left - GAP - WIDTH >= view.left ? cell.left - GAP - WIDTH : right
  return { top: `${Math.max(view.top, cell.top - GAP / 2)}px`, left: `${left}px` }
})

onMounted(() => {
  const el = input.value
  if (!el) return
  el.focus({ preventScroll: true })
  el.setSelectionRange(el.value.length, el.value.length)
})
useOutsidePress(box, () => void notes.save(false), () => !props.peek && !!notes.draft.value)

function onKey(e: KeyboardEvent): void {
  // the box's keys are its own: Enter is a new line, as in Excel's note
  e.stopPropagation()
  if (e.key !== 'Escape') return
  e.preventDefault()
  void notes.save()
}
</script>

<template>
  <div v-if="peek && shown && place" class="xs-note-box xs-note-peek" :style="place" role="note" :aria-label="m.note(shown.author)" data-xs-note-peek>
    <div v-if="shown.author" class="xs-note-author">{{ shown.author }}:</div>
    <div class="xs-note-text">{{ shown.text }}</div>
  </div>
  <div v-else-if="notes.draft.value && place" ref="box" class="xs-note-box" :style="place" data-xs-note-box>
    <div v-if="notes.draft.value.author" class="xs-note-author">{{ notes.draft.value.author }}:</div>
    <textarea
      ref="input"
      v-model="notes.draft.value.text"
      class="xs-note-text"
      :aria-label="m.noteEditor"
      spellcheck="true"
      data-xs-note-text
      @keydown="onKey"
    ></textarea>
  </div>
</template>

<style src="../links-notes.css"></style>
