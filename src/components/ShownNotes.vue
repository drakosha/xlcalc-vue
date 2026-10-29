<script setup lang="ts">
// Notes shown all the time, as Excel draws them: each in its box to the right of its cell, a little above it, with a line
// from the cell's top right corner (its indicator) to the box; read-only and transparent to the mouse, as the hover box is.
import { computed } from 'vue'
import { useSheetContext } from './context'

const ctx = useSheetContext()
const { v, views, notes } = ctx
const m = ctx.messages

/** How far the box stands to the right of its cell and above its top edge, as Excel places a shown note. */
const GAP = 12
const RISE = 8

const placed = computed(() =>
  notes.shown.value.flatMap(({ row, col, note }) => {
    const d = ctx.indexOf(row)
    if (d === undefined) return []
    const sp = views.spanOf(row, col)
    const cell = v.boxOf(d, col, d + (sp?.rs ?? 1) - 1, col + (sp?.cs ?? 1) - 1)
    const x1 = cell.left + cell.width
    const y1 = cell.top
    const left = x1 + GAP
    const top = Math.max(v.headH.value, cell.top - RISE)
    return [{ key: `${row}:${col}`, note, left, top, x1, y1 }]
  }),
)
/** The lines' layer covers the drawn notes. */
const layer = computed(() => {
  let w = 0
  let h = 0
  for (const p of placed.value) {
    w = Math.max(w, p.left + 1)
    h = Math.max(h, p.top + 1, p.y1 + 1)
  }
  return { width: `${w}px`, height: `${h}px` }
})
</script>

<template>
  <svg class="xs-note-lines" :style="layer" aria-hidden="true">
    <line v-for="p in placed" :key="p.key" :x1="p.x1" :y1="p.y1" :x2="p.left" :y2="p.top + RISE" :data-xs-note-line="p.key" />
  </svg>
  <div
    v-for="p in placed"
    :key="p.key"
    class="xs-note-box xs-note-peek xs-note-shown"
    :style="{ left: `${p.left}px`, top: `${p.top}px` }"
    role="note"
    :aria-label="m.note(p.note.author)"
    :data-xs-note-shown="p.key"
  >
    <div v-if="p.note.author" class="xs-note-author">{{ p.note.author }}:</div>
    <div class="xs-note-text">{{ p.note.text }}</div>
  </div>
</template>

<style src="../links-notes.css"></style>
