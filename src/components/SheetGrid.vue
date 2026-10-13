<script setup lang="ts">
// The grid: a real <table> with sticky letters and numbers, Excel's outline bars beside them, frozen panes, spacers for
// what is off screen, the marks over the cells (the copy's border, the fill handle) and the in-place field with its list of functions.
import { computed } from 'vue'
import { colLetters } from 'xlcalc/a1'
import { CfIcon } from './cf-icons'
import { useSheetContext } from './context'
import FunctionList from './FunctionList.vue'
import { IconFilter, IconSpinner } from './icons'
import OutlineCell from './OutlineCell.vue'
import OutlineLevels from './OutlineLevels.vue'
import SheetMarks from './SheetMarks.vue'
import ValidationDropdown from './ValidationDropdown.vue'
import ValidationPrompt from './ValidationPrompt.vue'

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
/** Drawn rows: the frozen ones (sticky), a spacer, the body's, a spacer; `d` is the display row, `r` the sheet's. */
const rows = computed(() => {
  const out: { key: string; d: number; r: number; frozen: boolean; pad: number }[] = []
  for (const d of v.frozenRowList.value) out.push({ key: `f${d}`, d, r: ctx.rowOf(d), frozen: true, pad: 0 })
  if (v.topPad.value) out.push({ key: 'tp', d: 0, r: 0, frozen: false, pad: v.topPad.value })
  for (const d of v.shownRows.value) out.push({ key: `b${d}`, d, r: ctx.rowOf(d), frozen: false, pad: 0 })
  if (v.bottomPad.value) out.push({ key: 'bp', d: 0, r: 0, frozen: false, pad: v.bottomPad.value })
  return out
})
const span = computed(() => (v.barW.value ? 1 : 0) + (v.showNums.value ? 1 : 0) + cols.value.length)
/** The outline bars' sizes for the sticky headers and numbers: left of the numbers, above the letters. */
const barVars = computed(() => ({ width: v.tableWidth.value + 'px', '--xs-outline-w': v.barW.value + 'px', '--xs-outline-h': v.barH.value + 'px' }))
const empty = computed(() => ctx.rowCount.value === 0 || ctx.colCount.value === 0)
const freezeRows = computed(() => ctx.sheet.value?.freeze?.rows ?? 0)
const freezeCols = computed(() => ctx.sheet.value?.freeze?.cols ?? 0)
const headLabel = (c: number) => ctx.sheet.value?.columns?.[c - 1]?.title ?? colLetters(c)
const colOn = (c: number) => !!sel.selRect.value && c >= sel.selRect.value.left && c <= sel.selRect.value.right
const rowOn = (r: number) => !!sel.selRect.value && r >= sel.selRect.value.top && r <= sel.selRect.value.bottom

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

/** The active cell (a merged area whole) in the scroller's content coordinates, for the dropdown and the input message of its rule. */
const activeBox = computed(() => {
  const a = ctx.selected.value
  if (!a) return null
  const d = ctx.indexOf(a.row)
  if (d === undefined) return null
  const sp = views.spanOf(a.row, a.col)
  const top = v.headH.value + (geo.rowOffsets.value[d - 1] ?? 0) + (d <= v.frozenRows.value ? ctx.scrollTop.value : 0)
  const left = v.leadW.value + (geo.offsets.value[a.col - 1] ?? 0) + (a.col <= v.frozenCols.value ? ctx.scrollLeft.value : 0)
  return { top, left, width: geo.spanWidth(a.col, sp?.cs ?? 1), height: geo.spanHeight(d, sp?.rs ?? 1) }
})
/** What of the grid is in view, in the same coordinates: a list goes over its cell where it has no room below. */
const viewBox = computed(() => ({
  top: ctx.scrollTop.value + v.headH.value,
  left: ctx.scrollLeft.value + v.leadW.value,
  width: Math.max(0, ctx.viewW.value - v.leadW.value),
  height: Math.max(0, ctx.viewH.value - v.headH.value),
}))
const vd = ctx.validation

/** A press on a cell: Excel's Point mode while a formula takes a reference there, else the selection's. */
function onCellDown(e: PointerEvent, r: number, c: number): void {
  if (!ctx.point.down(e, { kind: 'cells', row: r, col: c })) sel.onCellDown(e, r, c)
}
function onCellClick(e: MouseEvent, r: number, c: number): void {
  // The lift of a long press that opened the cells' menu keeps the selection, and so does a press that pointed.
  if (ctx.menu.afterPress() || ctx.point.took()) return
  // A double tap on a phone opens the field inside the tap, so the keyboard shows.
  if (sel.onCellTap(e, r, c)) ed.startEdit({ mode: 'edit' })
}
function onCellDbl(e: MouseEvent, r: number, c: number): void {
  if (ctx.point.took() || (e.target as HTMLElement).closest('a,button')) return
  sel.pickCell(r, c)
  ed.startEdit({ mode: 'edit' })
}
/** Funnels on the letters: when the source can filter and there is more than a header to filter; a finger gets one in the bar instead. */
const funnels = computed(() => ctx.filter.available.value && ctx.dataRows.value > 1 && !v.touch.value)
const filtered = (c: number) => !!ctx.filter.cols.value[c] || ctx.filter.sort.value?.col === c
function openFunnel(e: MouseEvent, c: number): void {
  const th = (e.currentTarget as HTMLElement).closest('th')
  void ctx.filter.openFunnel(c, th?.getBoundingClientRect())
}
// ── the menus on numbers, letters and cells: a right click, or a finger held still ──
/** The row number or column letter an event is on (not its edge or funnel). */
function headerOf(e: Event): { kind: 'rows' | 'columns'; line: number } | null {
  const el = e.target as HTMLElement
  const th = el.closest?.<HTMLElement>('th[data-row], th[data-col]')
  if (!th || el.closest('.xs-grip, .xs-funnel')) return null
  return th.dataset.row ? { kind: 'rows', line: Number(th.dataset.row) } : { kind: 'columns', line: Number(th.dataset.col) }
}
/** The cell an event is on, not a web link in it (the browser's menu has its own items for a link). */
function cellOf(e: Event): { row: number; col: number } | null {
  const el = e.target as HTMLElement
  const td = el.closest?.<HTMLElement>('td[data-cell]')
  if (!td || el.closest('a[href]')) return null
  const [row, col] = (td.dataset.cell ?? '').split(':').map(Number)
  return row && col ? { row, col } : null
}
function onContext(e: MouseEvent): void {
  const h = headerOf(e)
  const cell = h ? null : cellOf(e)
  if (!(h && ctx.menu.linesAvailable.value) && !cell) return
  e.preventDefault()
  ctx.menu.pressCancel()
  // A browser that opens its own menu on a long press sends this after ours opened.
  if (ctx.menu.afterPress()) return
  if (h) ctx.menu.openLines(h.kind, h.line, e.clientX, e.clientY)
  else ctx.menu.openCells(cell!.row, cell!.col, e.clientX, e.clientY)
}
function onDown(e: PointerEvent): void {
  const h = headerOf(e)
  const cell = h ? null : cellOf(e)
  // A cell's own press went to Point mode or the selection already; a letter or a number may point at its line.
  if (!cell && ctx.point.down(e, h && { kind: h.kind, row: h.kind === 'rows' ? h.line : 1, col: h.kind === 'columns' ? h.line : 1 })) return
  if (h && ctx.menu.linesAvailable.value) return ctx.menu.pressStart(e, (x, y) => ctx.menu.openLines(h.kind, h.line, x, y))
  if (cell) ctx.menu.pressStart(e, (x, y) => ctx.menu.openCells(cell.row, cell.col, x, y))
}
function onMove(e: PointerEvent): void {
  sel.onGridMove(e)
  ctx.point.move(e)
  ctx.menu.pressMove(e)
}
function onUp(): void {
  sel.onGridUp()
  ctx.point.up()
  ctx.menu.pressCancel()
}
/** A click on a letter or a number selects its line; the lift of a long press that opened the menu does not, nor a press that pointed. */
function pickLine(kind: 'rows' | 'columns', line: number, extend: boolean): void {
  if (ctx.menu.afterPress() || ctx.point.took()) return
  if (kind === 'rows') sel.pickRow(line, extend)
  else sel.pickColumn(line, extend)
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
    @keydown="ctx.onGridKey"
    @copy="ctx.clip.onCopy"
    @cut="ctx.clip.onCut"
    @paste="ctx.clip.onPaste"
    @contextmenu="onContext"
    @pointerdown="onDown"
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
              :class="['xs-head', k.c <= v.frozenCols.value ? 'xs-fz-col' : '', colOn(k.c) ? 'xs-on' : '', k.c === freezeCols ? 'xs-fz-edge-r' : '', filtered(k.c) ? 'xs-filtered' : '']"
              :style="k.c <= v.frozenCols.value ? { left: v.stickyLeft(k.c) } : undefined"
              :title="m.selectColumn"
              :data-col="k.c"
              @click="pickLine('columns', k.c, $event.shiftKey)"
            >
              <span class="xs-head-label"
                >{{ headLabel(k.c) }}<span v-if="ctx.filter.sort.value?.col === k.c" class="xs-sortmark">{{ ctx.filter.sort.value.dir === 'asc' ? '↑' : '↓' }}</span></span
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
          <tr v-else :style="{ height: geo.heightOf(row.d) + 'px' }">
            <OutlineCell v-if="v.barW.value" axis="rows" :line="row.r" :class="row.frozen ? 'xs-ol-fz' : ''" :style="row.frozen ? { top: v.stickyTop(row.d) } : undefined" />
            <th
              v-if="v.showNums.value"
              :class="['xs-num', row.frozen ? 'xs-fz-corner' : '', rowOn(row.r) ? 'xs-on' : '', row.d === freezeRows ? 'xs-fz-edge-b' : '']"
              :style="row.frozen ? { top: v.stickyTop(row.d) } : undefined"
              :title="m.selectRow"
              :data-row="row.r"
              @click="pickLine('rows', row.r, $event.shiftKey)"
            >
              <div class="xs-in xs-top" :style="{ height: Math.max(0, geo.heightOf(row.d) - 3) + 'px' }">
                <span>{{ row.r }}</span>
              </div>
              <span
                v-if="ctx.rowsResizable.value"
                class="xs-grip xs-grip-row"
                :title="m.resizeRow"
                @pointerdown="ctx.startRowResize($event, row.r)"
                @pointermove="ctx.onRowResizeMove"
                @pointerup="ctx.endRowResize"
                @pointercancel="ctx.endRowResize"
                @click.stop
                @dblclick.stop="ctx.onRowEdgeDouble(row.r)"
              ></span>
            </th>
            <template v-for="k in cols" :key="k.key">
              <td v-if="k.c === 0" :class="row.frozen ? 'xs-fz-row' : ''" :style="row.frozen ? { top: v.stickyTop(row.d) } : undefined"></td>
              <template v-else-if="!views.coveredByMerge(row.r, k.c)">
                <td
                  v-for="cv in [views.cellView(row.d, row.r, k.c, row.frozen)]"
                  :key="0"
                  v-bind="cv.attrs"
                  @pointerdown="onCellDown($event, row.r, k.c)"
                  @click="onCellClick($event, row.r, k.c)"
                  @dblclick="onCellDbl($event, row.r, k.c)"
                >
                  <div :class="cv.box.class" :style="cv.box.style">
                    <CfIcon v-if="cv.icon" :icon="cv.icon" />
                    <a
                      v-if="cv.kind === 'href'"
                      class="xs-link"
                      :style="cv.deco"
                      :href="cv.link!.href"
                      target="_blank"
                      rel="noopener noreferrer nofollow"
                      :title="cv.link!.tooltip ?? cv.link!.href"
                      @click="ctx.onLink($event, cv.link!, row.r, k.c)"
                      >{{ cv.text }}</a
                    >
                    <button
                      v-else-if="cv.kind === 'place' || cv.kind === 'host'"
                      type="button"
                      class="xs-link"
                      :style="cv.deco"
                      :title="cv.link!.tooltip ?? (cv.link!.location ? m.linkInBook(cv.link!.location) : m.linkOpen)"
                      @click.stop="ctx.onLink($event, cv.link!, row.r, k.c)"
                    >
                      {{ cv.text }}
                    </button>
                    <span v-else-if="cv.kind === 'dead'" class="xs-deadlink" :style="cv.deco" :title="cv.link!.tooltip ?? m.linkDead(cv.link!.href ?? cv.link!.target ?? '')">{{
                      cv.text
                    }}</span>
                    <div
                      v-else-if="cv.kind === 'overflow'"
                      class="xs-ov"
                      :style="{ left: -cv.ov!.left + 'px', width: `calc(100% + ${cv.ov!.left + cv.ov!.right}px)`, justifyContent: cv.ov!.justify }"
                    >
                      <span :style="{ ...cv.deco, background: cv.ov!.bg }">{{ cv.text }}</span>
                    </div>
                    <span v-else-if="cv.kind !== 'icon'" :style="cv.deco">{{ cv.text }}</span>
                  </div>
                </td>
              </template>
            </template>
          </tr>
        </template>
      </tbody>
    </table>
    <SheetMarks v-if="ctx.sheet.value && !empty" />
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
      @click="ed.ac.follow()"
      @blur="ed.onEditorBlur($event, 'cell')"
    ></textarea>
    <FunctionList
      v-if="editorRect && ed.ac.open.value"
      :rect="editorRect"
      :items="ed.ac.items.value"
      :active="ed.ac.index.value"
      :view="viewBox"
      :touch="v.touch.value"
      :messages="m"
      @highlight="(i) => (ed.ac.index.value = i)"
      @pick="(i) => ed.ac.accept(i)"
    />
    <template v-if="activeBox && !ctx.editor.value">
      <ValidationPrompt v-if="vd.prompt.value" :rect="activeBox" :prompt="vd.prompt.value" :view="viewBox" />
      <ValidationDropdown
        v-if="vd.list.value && ctx.editable.value"
        :rect="activeBox"
        :items="vd.list.value"
        :open="vd.open.value"
        :active="vd.index.value"
        :view="viewBox"
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
