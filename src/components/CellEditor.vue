<script setup lang="ts">
// The in-place field over its cell with its list of functions and tip of arguments: a component of its own, so a typed
// character redraws the field and not the grid.
import { computed } from 'vue'
import { useSheetContext } from './context'
import FunctionList from './FunctionList.vue'
import FunctionTip from './FunctionTip.vue'

const ctx = useSheetContext()
const { v, geo, views, ed } = ctx
const m = ctx.messages

/** The in-place field over its cell, in the scroller's content coordinates (a frozen cell's place follows the scroll). */
const editorRect = computed(() => {
  const e = ctx.editor.value
  if (!e || e.place !== 'cell' || e.sheet !== ctx.sheetName.value) return null
  const d = ctx.indexOf(e.row)
  if (d === undefined) return null
  const sp = views.spanOf(e.row, e.col)
  const frozen = d <= v.frozenRows.value || e.col <= v.frozenCols.value
  const top = v.headH.value + (geo.rowOffsets.value[d - 1] ?? 0) + (d <= v.frozenRows.value ? ctx.scrollTop.value : 0)
  const left = v.leadW.value + (geo.offsets.value[e.col - 1] ?? 0) + (e.col <= v.frozenCols.value ? ctx.scrollLeft.value : 0)
  const lines = e.text.split('\n')
  const longest = Math.max(...lines.map((l) => l.length))
  const room = Math.max(60, ctx.viewW.value - (left - ctx.scrollLeft.value) - 4)
  const width = Math.max(geo.spanWidth(e.col, sp?.cs ?? 1), Math.min(longest * 7.5 + 16, room))
  const height = Math.max(geo.spanHeight(d, sp?.rs ?? 1), lines.length * 17 + 5)
  return { top, left, width, height, frozen }
})
const editorBox = computed(() => {
  const r = editorRect.value
  return r && { top: `${r.top}px`, left: `${r.left}px`, width: `${r.width}px`, height: `${r.height}px`, zIndex: r.frozen ? '7' : '1' }
})
const setCellInput = (el: unknown) => (ctx.cellInput.value = (el as HTMLTextAreaElement | null) ?? null)
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
    v-if="editorRect && ed.ac.open.value"
    :rect="editorRect"
    :items="ed.ac.items.value"
    :active="ed.ac.index.value"
    :view="v.viewBox.value"
    :touch="v.touch.value"
    :messages="m"
    @highlight="(i) => (ed.ac.index.value = i)"
    @pick="(i) => ed.ac.accept(i)"
  />
  <FunctionTip v-if="editorRect && ed.tip.parts.value" :rect="editorRect" :parts="ed.tip.parts.value" :view="v.viewBox.value" />
</template>
