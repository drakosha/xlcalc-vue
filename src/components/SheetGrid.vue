<script setup lang="ts">
// The grid: a real <table> with sticky column letters and row numbers, Excel's outline bars next to them, frozen panes,
// spacers for what is off screen, the marks over the cells (the copy border, the fill handle) and the in-cell editor.
import { computed } from 'vue'
import { colLetters } from 'xlcalc/a1'
import CellEditor from './CellEditor.vue'
import { cellOfData } from '../core/selection'
import { useSheetContext } from './context'
import { refInto } from './dom'
import { IconFilter, IconSpinner } from './icons'
import OutlineCell from './OutlineCell.vue'
import OutlineLevels from './OutlineLevels.vue'
import SheetMarks from './SheetMarks.vue'
import SheetRow from './SheetRow.vue'
import ValidationDropdown from './ValidationDropdown.vue'
import ValidationPrompt from './ValidationPrompt.vue'

const ctx = useSheetContext()
const { v, geo, sel, views, ed, list } = ctx
const m = ctx.messages

/** Drawn columns: the frozen ones, a spacer for those scrolled away, the body's, a spacer for the rest (c = 0). */
const cols = computed(() => {
  const out: { key: string; c: number; w: number }[] = []
  for (const c of v.frozenColList.value) out.push({ key: `f${c}`, c, w: geo.widthOf(c) })
  if (v.leftPad.value) out.push({ key: 'lp', c: 0, w: v.leftPad.value })
  for (const c of v.shownCols.value) out.push({ key: `b${c}`, c, w: geo.widthOf(c) })
  if (v.rightPad.value) out.push({ key: 'rp', c: 0, w: v.rightPad.value })
  return out
})
/** The same columns for the rows, without widths: the same list while the columns are the same, so dragging a width does not redraw every row. */
const rowCols = computed((prev?: { key: string; c: number }[]) => {
  const next = cols.value.map(({ key, c }) => ({ key, c }))
  return prev && prev.length === next.length && prev.every((x, i) => x.key === next[i]!.key) ? prev : next
})
/** Drawn rows: the frozen ones (sticky), a spacer, the body's, a spacer; `d` is the display row, `r` the sheet row. */
const rows = computed(() => {
  const out: { key: string; d: number; r: number; frozen: boolean; pad: number }[] = []
  for (const d of v.frozenRowList.value) out.push({ key: `f${d}`, d, r: ctx.rowOf(d), frozen: true, pad: 0 })
  if (v.topPad.value) out.push({ key: 'tp', d: 0, r: 0, frozen: false, pad: v.topPad.value })
  for (const d of v.shownRows.value) out.push({ key: `b${d}`, d, r: ctx.rowOf(d), frozen: false, pad: 0 })
  if (v.bottomPad.value) out.push({ key: 'bp', d: 0, r: 0, frozen: false, pad: v.bottomPad.value })
  return out
})
const span = computed(() => (v.barW.value ? 1 : 0) + (v.showNums.value ? 1 : 0) + cols.value.length)
/** The outline bars' sizes for the sticky headers and numbers: left of the row numbers, above the column letters. */
const barVars = computed(() => ({ width: v.tableWidth.value + 'px', '--xs-outline-w': v.barW.value + 'px', '--xs-outline-h': v.barH.value + 'px' }))
const empty = computed(() => ctx.rowCount.value === 0 || ctx.colCount.value === 0)
const freezeCols = computed(() => ctx.sheet.value?.freeze?.cols ?? 0)
const headLabel = (c: number) => ctx.sheet.value?.columns?.[c - 1]?.title ?? colLetters(c)
const colOn = (c: number) => !!sel.selRect.value && c >= sel.selRect.value.left && c <= sel.selRect.value.right

/** The active cell (an entire merged range) in the scroller's content coordinates, for its rule's dropdown and input message. */
const activeBox = computed(() => {
  const a = ctx.selected.value
  if (!a) return null
  const d = ctx.indexOf(a.row)
  if (d === undefined) return null
  const sp = views.spanOf(a.row, a.col)
  const { top, left, width, height } = v.boxOf(d, a.col, d + (sp?.rs ?? 1) - 1, a.col + (sp?.cs ?? 1) - 1)
  return { top, left, width, height }
})
const vd = ctx.validation

// ── presses, clicks and double clicks on cells, handled on the scroller for every cell ──
/** A click on a cell: list mode opens what a cell leads to, and a double tap on a phone opens the editor. */
function onClick(e: MouseEvent): void {
  const at = cellUnder(e)
  // The release of a long press that opened the cell menu keeps the selection, and so does a press that inserted a reference.
  if (!at || ctx.menu.press.after() || ctx.point.took()) return
  // List mode: a plain click on a drill-down cell, or on a first-column cell, opens what it leads to.
  if (!ctx.editor.value && !e.shiftKey && !e.ctrlKey && !e.metaKey && list.click(at.row, at.col)) return
  // A double tap on a phone opens the editor within the tap, so the keyboard appears.
  if (sel.onCellTap(e, at.row, at.col)) ed.startEdit({ mode: 'edit' })
}
function onDouble(e: MouseEvent): void {
  const at = cellUnder(e)
  if (!at || ctx.point.took() || (e.target as HTMLElement).closest('a,button')) return
  sel.pickCell(at.row, at.col)
  ed.startEdit({ mode: 'edit' })
}
/** Funnels on the column letters: when the source can filter and there is more than a header to filter; touch input gets one in the formula bar instead. */
const funnels = computed(() => ctx.filter.available.value && ctx.dataRows.value > 1 && !v.touch.value)
const filtered = (c: number) => !!ctx.filter.cols.value[c] || ctx.filter.sort.value?.col === c
function openFunnel(e: MouseEvent, c: number): void {
  const th = (e.currentTarget as HTMLElement).closest('th')
  void ctx.filter.openFunnel(c, th?.getBoundingClientRect())
}
// ── the context menus on row numbers, column letters and cells: a right click, or a long press on a touch screen ──
/** The row number or column letter an event is on (not its edge or funnel). */
function headerOf(e: Event): { kind: 'rows' | 'columns'; line: number } | null {
  const el = e.target as HTMLElement
  const th = el.closest?.<HTMLElement>('th[data-row], th[data-col]')
  if (!th || el.closest('.xs-grip, .xs-funnel')) return null
  return th.dataset.row ? { kind: 'rows', line: Number(th.dataset.row) } : { kind: 'columns', line: Number(th.dataset.col) }
}
/** The cell an event is on, including a link in it. */
function cellUnder(e: Event): { row: number; col: number } | null {
  const td = (e.target as HTMLElement).closest?.<HTMLElement>('td[data-cell]')
  return td ? cellOfData(td.dataset.cell) : null
}
/** The same, but not a web link in it (the browser's menu has its own items for a link). */
const cellOf = (e: Event) => ((e.target as HTMLElement).closest?.('a[href]') ? null : cellUnder(e))
function onContext(e: MouseEvent): void {
  const h = headerOf(e)
  const cell = h ? null : cellOf(e)
  if (!(h && ctx.menu.linesAvailable.value) && !cell) return
  if (!ctx.menu.press.context(e)) return
  if (h) ctx.menu.openLines(h.kind, h.line, e.clientX, e.clientY)
  else ctx.menu.openCells(cell!.row, cell!.col, e.clientX, e.clientY)
}
function onDown(e: PointerEvent): void {
  const h = headerOf(e)
  const under = h ? null : cellUnder(e)
  // A press on a cell: Excel's Point mode while a formula accepts a reference there, otherwise the selection.
  if (under && !ctx.point.down(e, { kind: 'cells', row: under.row, col: under.col })) sel.onCellDown(e, under.row, under.col)
  const cell = under && cellOf(e)
  // A column letter or a row number may insert a reference to its column or row.
  if (!cell && ctx.point.down(e, h && { kind: h.kind, row: h.kind === 'rows' ? h.line : 1, col: h.kind === 'columns' ? h.line : 1 })) return
  if (h && ctx.menu.linesAvailable.value) return ctx.menu.press.start(e, (x, y) => ctx.menu.openLines(h.kind, h.line, x, y))
  if (cell) ctx.menu.press.start(e, (x, y) => ctx.menu.openCells(cell.row, cell.col, x, y))
}
function onMove(e: PointerEvent): void {
  sel.onGridMove(e)
  ctx.point.move(e)
  ctx.menu.press.move(e)
}
function onUp(): void {
  sel.onGridUp()
  ctx.point.up()
  ctx.menu.press.cancel()
}
/** A click on a column letter or a row number selects its column or row, and a sortable title in list mode sorts; the release of a long press that opened the menu does not, nor does a press that inserted a reference. */
function pickLine(kind: 'rows' | 'columns', line: number, extend: boolean): void {
  if (ctx.menu.press.after() || ctx.point.took()) return
  if (kind === 'columns' && !extend && list.sortable(line)) return list.toggleSort(line)
  if (kind === 'rows') sel.pickRow(line, extend)
  else sel.pickColumn(line, extend)
}
const setScroller = refInto(ctx.scroller)
</script>

<template>
  <div
    :ref="setScroller"
    class="xs-scroller"
    tabindex="0"
    @scroll.passive="ctx.onScroll"
    @keydown="ctx.onGridKey"
    @copy="ctx.clip.onCopy"
    @cut="ctx.clip.onCut"
    @paste="ctx.clip.onPaste"
    @contextmenu="onContext"
    @pointerdown="onDown"
    @click="onClick"
    @dblclick="onDouble"
    @mousedown="ctx.point.onMouseDown"
    @pointermove="onMove"
    @pointerup="onUp"
    @pointercancel="onUp"
  >
    <p v-if="empty && ctx.sheet.value" class="xs-empty">{{ m.empty }}</p>
    <table v-else-if="ctx.sheet.value" class="xs-table" :style="barVars">
      <colgroup>
        <col v-if="v.barW.value" :style="{ width: v.barW.value + 'px' }" />
        <col v-if="v.showNums.value" :style="{ width: v.numW.value + 'px' }" />
        <col v-for="k in cols" :key="k.key" :style="{ width: k.w + 'px' }" />
      </colgroup>
      <thead>
        <tr v-if="v.barH.value" class="xs-ol-head" :style="{ height: v.barH.value + 'px' }">
          <th v-if="v.barW.value && v.showNums.value" class="xs-ol-corner"></th>
          <OutlineLevels v-if="v.barW.value || v.showNums.value" axis="columns" :class="v.showNums.value ? 'xs-ol-over-nums' : ''" />
          <template v-for="k in cols" :key="k.key">
            <th v-if="k.c === 0" class="xs-ol xs-ol-columns"></th>
            <OutlineCell
              v-else
              axis="columns"
              :line="k.c"
              :class="k.c <= v.frozenCols.value ? 'xs-ol-fz' : ''"
              :style="k.c <= v.frozenCols.value ? { left: v.stickyLeft(k.c) } : undefined"
            />
          </template>
        </tr>
        <tr>
          <OutlineLevels v-if="v.barW.value" axis="rows" />
          <th v-if="v.showNums.value" class="xs-corner" :title="m.selectAll" @click="sel.selectAll()"></th>
          <template v-for="k in cols" :key="k.key">
            <th v-if="k.c === 0" class="xs-head"></th>
            <th
              v-else
              :class="[
                'xs-head',
                k.c <= v.frozenCols.value ? 'xs-fz-col' : '',
                colOn(k.c) ? 'xs-on' : '',
                k.c === freezeCols ? 'xs-fz-edge-r' : '',
                filtered(k.c) ? 'xs-filtered' : '',
                list.sortable(k.c) ? 'xs-sortable' : '',
              ]"
              :style="k.c <= v.frozenCols.value ? { left: v.stickyLeft(k.c) } : undefined"
              :title="list.sortable(k.c) ? m.sortColumn : m.selectColumn"
              :aria-sort="list.sortOf(k.c) === 'asc' ? 'ascending' : list.sortOf(k.c) === 'desc' ? 'descending' : undefined"
              :data-col="k.c"
              @click="pickLine('columns', k.c, $event.shiftKey)"
            >
              <span class="xs-head-label" :style="list.alignOf(k.c) ? { textAlign: list.alignOf(k.c) } : undefined"
                >{{ headLabel(k.c) }}<span v-if="ctx.filter.sort.value?.col === k.c" class="xs-sortmark">{{ ctx.filter.sort.value.dir === 'asc' ? '↑' : '↓' }}</span
                ><span v-else-if="list.sortOf(k.c)" class="xs-sortmark xs-list-sort">{{ list.sortOf(k.c) === 'asc' ? '↑' : '↓' }}</span
                ><span v-else-if="list.sortable(k.c)" class="xs-sortmark xs-sorthint" aria-hidden="true">↕</span></span
              >
              <button
                v-if="funnels"
                type="button"
                :class="['xs-funnel', ctx.filter.cols.value[k.c] ? 'xs-on' : '', ctx.filter.open.value === k.c ? 'xs-open' : '']"
                :title="ctx.filter.cols.value[k.c] ? m.filterOn : m.filter"
                @click.stop="openFunnel($event, k.c)"
              >
                <IconFilter />
              </button>
              <span
                class="xs-grip"
                :title="ctx.fitsOnEdge.value ? m.resizeFit : m.resize"
                @pointerdown="ctx.startResize($event, k.c)"
                @pointermove="ctx.onResizeMove"
                @pointerup="ctx.endResize"
                @pointercancel="ctx.endResize"
                @click.stop
                @dblclick.stop="ctx.onEdgeDouble(k.c)"
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
          <SheetRow v-else :d="row.d" :r="row.r" :frozen="row.frozen" :cols="rowCols" @pick="(r, extend) => pickLine('rows', r, extend)" />
        </template>
      </tbody>
      <tfoot v-if="list.totals.value">
        <tr>
          <th v-if="v.barW.value" class="xs-total xs-total-bar"></th>
          <th v-if="v.showNums.value" class="xs-total xs-total-num"></th>
          <template v-for="k in cols" :key="k.key">
            <td v-if="k.c === 0" class="xs-total"></td>
            <td
              v-else
              :class="['xs-total', k.c <= v.frozenCols.value ? 'xs-fz-col' : '', k.c === freezeCols ? 'xs-fz-edge-r' : '']"
              :style="{ ...(k.c <= v.frozenCols.value ? { left: v.stickyLeft(k.c) } : {}), ...(list.alignOf(k.c) ? { textAlign: list.alignOf(k.c) } : {}) }"
              :data-col="k.c"
              :title="list.totals.value[k.c - 1] ?? ''"
            >
              {{ list.totals.value[k.c - 1] ?? '' }}
            </td>
          </template>
        </tr>
      </tfoot>
    </table>
    <SheetMarks v-if="ctx.sheet.value && !empty" />
    <CellEditor v-if="ctx.editable.value" />
    <template v-if="activeBox && !ctx.editor.value">
      <ValidationPrompt v-if="vd.prompt.value" :rect="activeBox" :prompt="vd.prompt.value" :view="v.viewBox.value" />
      <ValidationDropdown
        v-if="vd.list.value && ctx.editable.value"
        :rect="activeBox"
        :items="vd.list.value"
        :open="vd.open.value"
        :active="vd.index.value"
        :view="v.viewBox.value"
        :touch="v.touch.value"
        :messages="m"
        @toggle="vd.toggle()"
        @pick="(i) => vd.choose(i)"
        @hover="(i) => (vd.index.value = i)"
        @close="vd.close()"
      />
    </template>
    <div v-if="ctx.loading.value" class="xs-loading">
      <IconSpinner class="xs-spin" />
      {{ m.loading }}
    </div>
  </div>
</template>
