<script setup lang="ts">
// Excel's Watch Window: a window floating over the sheet within its box, moved by its title bar and resized by its corner,
// with Add Watch, Delete Watch and the watched cells in Excel's columns; a double click (or Enter) goes to a cell, Delete removes the chosen rows.
import { ref } from 'vue'
import { onMac } from '../core/keys'
import { useSheetContext } from './context'
import { IconClose } from './icons'

const ctx = useSheetContext()
const m = ctx.messages
const wt = ctx.watches

const box = ref<HTMLElement | null>(null)
/** Where the window was moved to, within the sheet's box; null: its place in the CSS (the top right corner). */
const pos = ref<{ left: number; top: number } | null>(null)
let drag: { x: number; y: number; left: number; top: number } | null = null
function onHeadDown(e: PointerEvent): void {
  const el = box.value
  if (!el || e.button !== 0 || (e.target as HTMLElement).closest('button')) return
  drag = { x: e.clientX, y: e.clientY, left: el.offsetLeft, top: el.offsetTop }
  ;(e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId)
  e.preventDefault()
}
function onHeadMove(e: PointerEvent): void {
  const el = box.value
  const parent = el?.offsetParent as HTMLElement | null
  if (!drag || !el) return
  const maxLeft = Math.max(0, (parent?.clientWidth ?? 0) - el.offsetWidth)
  const maxTop = Math.max(0, (parent?.clientHeight ?? 0) - el.offsetHeight)
  pos.value = { left: Math.min(maxLeft, Math.max(0, drag.left + e.clientX - drag.x)), top: Math.min(maxTop, Math.max(0, drag.top + e.clientY - drag.y)) }
}
const onHeadUp = () => void (drag = null)

function onRow(e: MouseEvent, i: number): void {
  wt.choose(i, { add: onMac() ? e.metaKey : e.ctrlKey, range: e.shiftKey })
}
/** The list's keys: the arrows move the chosen row, Enter goes to it, Delete and Backspace remove the chosen rows. */
function onKey(e: KeyboardEvent): void {
  const n = wt.list.value.length
  const at = Math.max(-1, ...wt.chosen.value)
  if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && n) {
    e.preventDefault()
    wt.choose(Math.min(n - 1, Math.max(0, at + (e.key === 'ArrowDown' ? 1 : -1))), { range: e.shiftKey })
  } else if (e.key === 'Enter' && at >= 0) {
    e.preventDefault()
    void wt.go(at)
  } else if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault()
    wt.remove()
  }
}
function close(): void {
  wt.setOpen(false)
  ctx.focusGrid()
}
</script>

<template>
  <section ref="box" class="xs-watch" role="dialog" :aria-label="m.watchTitle" :style="pos ? { left: `${pos.left}px`, top: `${pos.top}px`, right: 'auto', bottom: 'auto' } : undefined" data-xs-watch>
    <div class="xs-watch-head" @pointerdown="onHeadDown" @pointermove="onHeadMove" @pointerup="onHeadUp" @pointercancel="onHeadUp">
      <span class="xs-watch-title">{{ m.watchTitle }}</span>
      <button type="button" class="xs-tool" :title="m.close" :aria-label="m.close" data-xs-watch-close @click="close">
        <IconClose />
      </button>
    </div>
    <div class="xs-watch-tools">
      <button type="button" class="xs-button" :disabled="!wt.canAdd.value" data-xs-watch-add @click="wt.add()">{{ m.watchAdd }}</button>
      <button type="button" class="xs-button" :disabled="!wt.chosen.value.size" data-xs-watch-delete @click="wt.remove()">{{ m.watchDelete }}</button>
    </div>
    <div class="xs-watch-body" tabindex="0" data-xs-watch-list @keydown="onKey">
      <table class="xs-watch-table">
        <thead>
          <tr>
            <th>{{ m.watchBook }}</th>
            <th>{{ m.watchSheet }}</th>
            <th>{{ m.watchName }}</th>
            <th>{{ m.watchCell }}</th>
            <th>{{ m.watchValue }}</th>
            <th>{{ m.watchFormula }}</th>
          </tr>
        </thead>
        <tbody>
          <tr
            v-for="(w, i) in wt.list.value"
            :key="`${w.sheet}!${w.cell}`"
            :class="{ 'xs-on': wt.chosen.value.has(i) }"
            :aria-selected="wt.chosen.value.has(i)"
            :data-xs-watch-row="`${w.sheet}!${w.cell}`"
            @click="onRow($event, i)"
            @dblclick="wt.go(i)"
          >
            <!-- the workbook's name and the cell's defined name are not known to the sheet yet -->
            <td></td>
            <td :title="w.sheet">{{ w.sheet }}</td>
            <td></td>
            <td>{{ w.cell }}</td>
            <td :title="wt.valueOf(w)?.value" data-xs-watch-value>{{ wt.valueOf(w)?.value ?? '' }}</td>
            <td :title="wt.valueOf(w)?.formula" data-xs-watch-formula>{{ wt.valueOf(w)?.formula ?? '' }}</td>
          </tr>
        </tbody>
      </table>
    </div>
  </section>
</template>

<style src="../view.css"></style>
