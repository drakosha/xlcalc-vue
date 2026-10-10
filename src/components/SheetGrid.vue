<script setup lang="ts">
// The grid: a real <table> with sticky letters and numbers, frozen panes, spacers for what is off screen, and the in-place field.
import { computed } from 'vue'
import { colLabel } from '../core/a1'
import { HEAD_H } from '../core/viewport'
import { useSheetContext } from './context'
import { IconSpinner } from './icons'

const ctx = useSheetContext()
const { v, geo, sel, views, ed } = ctx
const m = ctx.messages

/** Drawn columns: the frozen ones, a spacer for the scrolled-away ones, the body's, a spacer for the rest (c = 0). */
const cols = computed(() => {
  const out: { key: string; c: number; w: number }[] = []
  for (const c of v.frozenColList.value) out.push({ key: `f${c}`, c, w: geo.widthOf(c) })
  if (v.leftPad.value) out.push({ key: 'lp', c: 0, w: v.leftPad.value })
  for (const c of v.shownCols.value) out.push({ key: `b${c}`, c, w: geo.widthOf(c) })
  if (v.rightPad.value) out.push({ key: 'rp', c: 0, w: v.rightPad.value })
  return out
})
/** Drawn rows: the frozen ones (sticky), a spacer, the body's, a spacer; `d` is the display row. */
const rows = computed(() => {
  const out: { key: string; d: number; frozen: boolean; pad: number }[] = []
  for (const d of v.frozenRowList.value) out.push({ key: `f${d}`, d, frozen: true, pad: 0 })
  if (v.topPad.value) out.push({ key: 'tp', d: 0, frozen: false, pad: v.topPad.value })
  for (const d of v.shownRows.value) out.push({ key: `b${d}`, d, frozen: false, pad: 0 })
  if (v.bottomPad.value) out.push({ key: 'bp', d: 0, frozen: false, pad: v.bottomPad.value })
  return out
})
const span = computed(() => (v.showNums.value ? 1 : 0) + cols.value.length)
const empty = computed(() => ctx.rowCount.value === 0 || ctx.colCount.value === 0)
const freezeRows = computed(() => ctx.sheet.value?.freeze?.rows ?? 0)
const freezeCols = computed(() => ctx.sheet.value?.freeze?.cols ?? 0)
const headLabel = (c: number) => ctx.sheet.value?.columns?.[c - 1]?.title ?? colLabel(c)
const colOn = (c: number) => !!sel.selRect.value && c >= sel.selRect.value.c1 && c <= sel.selRect.value.c2
const rowOn = (r: number) => !!sel.selRect.value && r >= sel.selRect.value.r1 && r <= sel.selRect.value.r2

/** The in-place field over its cell, in the scroller's content coordinates (a frozen cell's place follows the scroll). */
const editorBox = computed(() => {
  const e = ctx.editor.value
  if (!e || e.place !== 'cell' || e.sheet !== ctx.sheetName.value) return null
  const sp = views.spanOf(e.row, e.col)
  const frozen = e.row <= v.frozenRows.value || e.col <= v.frozenCols.value
  const top = HEAD_H + (geo.rowOffsets.value[e.row - 1] ?? 0) + (e.row <= v.frozenRows.value ? ctx.scrollTop.value : 0)
  const left = v.numW.value + (geo.offsets.value[e.col - 1] ?? 0) + (e.col <= v.frozenCols.value ? ctx.scrollLeft.value : 0)
  const lines = e.text.split('\n')
  const longest = Math.max(...lines.map((l) => l.length))
  const room = Math.max(60, ctx.viewW.value - (left - ctx.scrollLeft.value) - 4)
  const width = Math.max(geo.spanWidth(e.col, sp?.cs ?? 1), Math.min(longest * 7.5 + 16, room))
  const height = Math.max(geo.spanHeight(e.row, sp?.rs ?? 1), lines.length * 17 + 5)
  return { top: `${top}px`, left: `${left}px`, width: `${width}px`, height: `${height}px`, zIndex: frozen ? '7' : '1' }
})

function onCellClick(e: MouseEvent, r: number, c: number): void {
  // A double tap on a phone opens the field inside the tap, so the keyboard shows.
  if (sel.onCellTap(e, r, c)) ed.startEdit({ mode: 'edit' })
}
function onCellDbl(e: MouseEvent, r: number, c: number): void {
  if ((e.target as HTMLElement).closest('a,button')) return
  sel.pickCell(r, c)
  ed.startEdit({ mode: 'edit' })
}
const setCellInput = (el: unknown) => (ctx.cellInput.value = (el as HTMLTextAreaElement | null) ?? null)
const setScroller = (el: unknown) => (ctx.scroller.value = (el as HTMLElement | null) ?? null)
</script>

<template>
  <div
    :ref="setScroller"
    class="xs-scroller"
    tabindex="0"
    @scroll.passive="ctx.onScroll"
    @keydown="ed.onGridKey"
    @copy="ctx.totals.onCopy"
    @paste="ed.onPaste"
    @pointermove="sel.onGridMove"
    @pointerup="sel.onGridUp"
    @pointercancel="sel.onGridUp"
  >
    <p v-if="empty && ctx.sheet.value" class="xs-empty">{{ m.empty }}</p>
    <table v-else-if="ctx.sheet.value" class="xs-table" :style="{ width: v.tableWidth.value + 'px' }">
      <colgroup>
        <col v-if="v.showNums.value" :style="{ width: v.numW.value + 'px' }" />
        <col v-for="k in cols" :key="k.key" :style="{ width: k.w + 'px' }" />
      </colgroup>
      <thead>
        <tr>
          <th v-if="v.showNums.value" class="xs-corner" :title="m.selectAll" @click="sel.selectAll()"></th>
          <template v-for="k in cols" :key="k.key">
            <th v-if="k.c === 0" class="xs-head"></th>
            <th
              v-else
              :class="['xs-head', k.c <= v.frozenCols.value ? 'xs-fz-col' : '', colOn(k.c) ? 'xs-on' : '', k.c === freezeCols ? 'xs-fz-edge-r' : '']"
              :style="k.c <= v.frozenCols.value ? { left: v.stickyLeft(k.c) } : undefined"
              :title="m.selectColumn"
              @click="sel.pickColumn(k.c, $event.shiftKey)"
            >
              {{ headLabel(k.c) }}
              <span
                class="xs-grip"
                :title="m.resize"
                @pointerdown="ctx.startResize($event, k.c)"
                @pointermove="ctx.onResizeMove"
                @pointerup="ctx.endResize"
                @pointercancel="ctx.endResize"
                @click.stop
                @dblclick.stop="ctx.resetWidth(k.c)"
              ></span>
            </th>
          </template>
        </tr>
      </thead>
      <tbody>
        <template v-for="row in rows" :key="row.key">
          <tr v-if="row.pad" class="xs-spacer" :style="{ height: row.pad + 'px' }">
            <td :colspan="span"></td>
          </tr>
          <tr v-else :style="{ height: geo.heightOf(row.d) + 'px' }">
            <th
              v-if="v.showNums.value"
              :class="['xs-num', row.frozen ? 'xs-fz-corner' : '', rowOn(row.d) ? 'xs-on' : '', row.d === freezeRows ? 'xs-fz-edge-b' : '']"
              :style="row.frozen ? { top: v.stickyTop(row.d) } : undefined"
              :title="m.selectRow"
              @click="sel.pickRow(row.d, $event.shiftKey)"
            >
              <div class="xs-in xs-top" :style="{ height: Math.max(0, geo.heightOf(row.d) - 3) + 'px' }">
                <span>{{ row.d }}</span>
              </div>
            </th>
            <template v-for="k in cols" :key="k.key">
              <td v-if="k.c === 0" :class="row.frozen ? 'xs-fz-row' : ''" :style="row.frozen ? { top: v.stickyTop(row.d) } : undefined"></td>
              <template v-else-if="!views.coveredByMerge(row.d, k.c)">
                <td
                  v-for="cv in [views.cellView(row.d, row.d, k.c, row.frozen)]"
                  :key="0"
                  v-bind="cv.attrs"
                  @pointerdown="sel.onCellDown($event, row.d, k.c)"
                  @click="onCellClick($event, row.d, k.c)"
                  @dblclick="onCellDbl($event, row.d, k.c)"
                >
                  <div :class="cv.box.class" :style="cv.box.style">
                    <a
                      v-if="cv.kind === 'href'"
                      class="xs-link"
                      :href="cv.link!.href"
                      target="_blank"
                      rel="noopener noreferrer nofollow"
                      :title="cv.link!.tooltip ?? cv.link!.href"
                      @click="ctx.onLink($event, cv.link!, row.d, k.c)"
                      >{{ cv.text }}</a
                    >
                    <button
                      v-else-if="cv.kind === 'place' || cv.kind === 'host'"
                      type="button"
                      class="xs-link"
                      :title="cv.link!.tooltip ?? (cv.link!.location ? m.linkInBook(cv.link!.location) : m.linkOpen)"
                      @click.stop="ctx.onLink($event, cv.link!, row.d, k.c)"
                    >
                      {{ cv.text }}
                    </button>
                    <span v-else-if="cv.kind === 'dead'" class="xs-deadlink" :title="cv.link!.tooltip ?? m.linkDead(cv.link!.href ?? cv.link!.target ?? '')">{{
                      cv.text
                    }}</span>
                    <div
                      v-else-if="cv.kind === 'overflow'"
                      class="xs-ov"
                      :style="{ left: -cv.ov!.left + 'px', width: `calc(100% + ${cv.ov!.left + cv.ov!.right}px)`, justifyContent: cv.ov!.justify }"
                    >
                      <span>{{ cv.text }}</span>
                    </div>
                    <span v-else>{{ cv.text }}</span>
                  </div>
                </td>
              </template>
            </template>
          </tr>
        </template>
      </tbody>
    </table>
    <textarea
      v-if="ctx.editable.value"
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
      @blur="ed.onEditorBlur($event, 'cell')"
    ></textarea>
    <div v-if="ctx.loading.value" class="xs-loading">
      <IconSpinner class="xs-spin" />
      {{ m.loading }}
    </div>
  </div>
</template>
