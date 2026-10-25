<script setup lang="ts">
// The in-cell editor over its cell, with its function list and argument tooltip: a separate component, so a typed
// character redraws the editor and not the grid. The editor is drawn at the sheet's zoom; the list and the tooltip are not, as in Excel.
import { computed } from 'vue'
import { useSheetContext } from './context'
import { refInto } from './dom'
import FunctionList from './FunctionList.vue'
import FunctionTip from './FunctionTip.vue'

const ctx = useSheetContext()
const { v, views, ed } = ctx
const m = ctx.messages

/** The in-cell editor over its cell, in the scroller's content coordinates (a frozen cell's position follows the scroll). */
const editorRect = computed(() => {
  const e = ctx.editor.value
  if (!e || e.place !== 'cell' || e.sheet !== ctx.sheetName.value) return null
  const d = ctx.indexOf(e.row)
  if (d === undefined) return null
  // A merged range is drawn as one cell only on one side of the freeze edge, so its box follows the scroll as its first cell does.
  const sp = views.spanOf(e.row, e.col)
  const cell = v.boxOf(d, e.col, d + (sp?.rs ?? 1) - 1, e.col + (sp?.cs ?? 1) - 1)
  const lines = e.text.split('\n')
  const longest = Math.max(...lines.map((l) => l.length))
  const room = Math.max(60, ctx.viewW.value - (cell.left - ctx.scrollLeft.value) - 4)
  const width = Math.max(cell.width, Math.min(longest * 7.5 + 16, room))
  const height = Math.max(cell.height, lines.length * 17 + 5)
  return { top: cell.top, left: cell.left, width, height, frozen: cell.frozen }
})
const editorBox = computed(() => {
  const r = editorRect.value
  const z = ctx.zoom.value
  return r && { top: `${r.top}px`, left: `${r.left}px`, width: `${r.width}px`, height: `${r.height}px`, zIndex: r.frozen ? '7' : '1', ...(z !== 1 ? { zoom: String(z) } : {}) }
})
/** A box in the grid's unzoomed pixels on the scroller's zoomed content, for the popups drawn at their own size. */
const zoomed = <T extends { top: number; left: number; width: number; height: number }>(b: T): T => {
  const z = ctx.zoom.value
  return z === 1 ? b : { ...b, top: b.top * z, left: b.left * z, width: b.width * z, height: b.height * z }
}
const popupRect = computed(() => editorRect.value && zoomed(editorRect.value))
const popupView = computed(() => zoomed(v.viewBox.value))
const setCellInput = refInto(ctx.cellInput)
</script>

<template>
  <textarea
    :ref="setCellInput"
    :class="['xs-editor', editorBox ? '' : 'xs-editor-idle']"
    :style="editorBox ?? undefined"
    :value="ctx.editor.value?.place === 'cell' ? ctx.editor.value.text : ''"
    rows="1"
    spellcheck="false"
    autocomplete="off"
    autocapitalize="off"
    :tabindex="editorBox ? 0 : -1"
    :aria-hidden="editorBox ? undefined : 'true'"
    @input="ed.onEditorInput"
    @keydown="ed.onEditorKey"
    @click="ed.onEditorClick()"
    @blur="ed.onEditorBlur($event, 'cell')"
  ></textarea>
  <FunctionList
    v-if="popupRect && ed.ac.open.value"
    :rect="popupRect"
    :items="ed.ac.items.value"
    :active="ed.ac.index.value"
    :view="popupView"
    :touch="v.touch.value"
    :messages="m"
    @highlight="(i) => (ed.ac.index.value = i)"
    @pick="(i) => ed.ac.accept(i)"
  />
  <FunctionTip v-if="popupRect && ed.tip.parts.value" :rect="popupRect" :parts="ed.tip.parts.value" :view="popupView" />
</template>
