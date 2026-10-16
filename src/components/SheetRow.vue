<script setup lang="ts">
// A drawn row of the grid: its outline bar's cell, its number and its cells. A component of its own, so a change redraws
// only the rows it touches: the cells read the selection, the highlights, the hits, the flash and the field by their row.
import { computed, type ComputedRef } from 'vue'
import { sameRowState, type CellView, type RowState } from '../core/cell-view'
import CellDraft from './CellDraft.vue'
import { CfIcon } from './cf-icons'
import { useSheetContext } from './context'
import OutlineCell from './OutlineCell.vue'

const props = defineProps<{
  /** The display row and the sheet's. */
  d: number
  r: number
  frozen: boolean
  /** Drawn columns, a spacer's with c = 0. */
  cols: readonly { key: string; c: number }[]
}>()
const emit = defineEmits<{
  /** A click on the row's number: the row picked, extended with Shift. */
  pick: [row: number, extend: boolean]
}>()
const ctx = useSheetContext()
const { v, views } = ctx
const m = ctx.messages
const freezeRows = computed(() => ctx.sheet.value?.freeze?.rows ?? 0)

/** The often changing state over this row, the same object while it is the same: the row is drawn again only when it changes. */
const state = computed((prev?: RowState) => {
  const next = views.rowState(props.d, props.r)
  return prev && sameRowState(prev, next) ? prev : next
})
/**
 * Each drawn cell's view as a computed of its own (null where a merge covers the cell): a column scrolled in or out
 * leaves the other cells' views as they were, and their memo leaves their DOM alone.
 */
const viewOfCol = new Map<number, ComputedRef<CellView | null>>()
function viewOf(c: number): CellView | null {
  let view = viewOfCol.get(c)
  if (!view) viewOfCol.set(c, (view = computed(() => (views.coveredByMerge(props.r, c) ? null : views.cellView(props.d, props.r, c, props.frozen, state.value)))))
  return view.value
}
/** The row's cells by column: a spacer's empty cell (no view), else the cell's view; cells a merge covers are left out. */
const cells = computed(() => {
  const out: { key: string; c: number; cv: CellView | null }[] = []
  const drawn = new Set<number>()
  for (const k of props.cols) {
    if (k.c === 0) out.push({ key: k.key, c: 0, cv: null })
    else {
      drawn.add(k.c)
      const cv = viewOf(k.c)
      if (cv) out.push({ key: k.key, c: k.c, cv })
    }
  }
  // the views of columns no longer drawn go
  for (const c of viewOfCol.keys()) if (!drawn.has(c)) viewOfCol.delete(c)
  return out
})
/** A spacer's cell: sticky in a frozen row. */
const spacer = computed(() => (props.frozen ? { class: 'xs-fz-row', style: { top: v.stickyTop(props.d) } } : {}))
</script>

<template>
  <tr :style="{ height: state.height + 'px' }">
    <OutlineCell v-if="v.barW.value" axis="rows" :line="r" :class="frozen ? 'xs-ol-fz' : ''" :style="frozen ? { top: v.stickyTop(d) } : undefined" />
    <th
      v-if="v.showNums.value"
      :class="['xs-num', frozen ? 'xs-fz-corner' : '', state.sel ? 'xs-on' : '', d === freezeRows ? 'xs-fz-edge-b' : '']"
      :style="frozen ? { top: v.stickyTop(d) } : undefined"
      :title="m.selectRow"
      :data-row="r"
      @click="emit('pick', r, $event.shiftKey)"
    >
      <div class="xs-in xs-top" :style="{ height: Math.max(0, state.height - 3) + 'px' }">
        <span>{{ r }}</span>
      </div>
      <span
        v-if="ctx.rowsResizable.value"
        class="xs-grip xs-grip-row"
        :title="m.resizeRow"
        @pointerdown="ctx.startRowResize($event, r)"
        @pointermove="ctx.onRowResizeMove"
        @pointerup="ctx.endRowResize"
        @pointercancel="ctx.endRowResize"
        @click.stop
        @dblclick.stop="ctx.onRowEdgeDouble(r)"
      ></span>
    </th>
    <!-- a cell whose view is the same object as before is not patched: its view stands for everything it draws -->
    <td v-for="{ key, c, cv } in cells" :key="key" v-memo="[cv, cv ? m : spacer]" v-bind="cv ? cv.attrs : spacer">
      <template v-if="cv">
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
            @click="ctx.onLink($event, cv.link!, r, c)"
            ><CellDraft v-if="cv.draft" /><template v-else>{{ cv.text }}</template></a
          >
          <button
            v-else-if="cv.kind === 'place' || cv.kind === 'host'"
            type="button"
            class="xs-link"
            :style="cv.deco"
            :title="cv.link!.tooltip ?? (cv.link!.location ? m.linkInBook(cv.link!.location) : m.linkOpen)"
            @click.stop="ctx.onLink($event, cv.link!, r, c)"
          >
            <CellDraft v-if="cv.draft" /><template v-else>{{ cv.text }}</template>
          </button>
          <span v-else-if="cv.kind === 'dead'" class="xs-deadlink" :style="cv.deco" :title="cv.link!.tooltip ?? m.linkDead(cv.link!.href ?? cv.link!.target ?? '')"
            ><CellDraft v-if="cv.draft" /><template v-else>{{ cv.text }}</template></span
          >
          <div
            v-else-if="cv.kind === 'overflow'"
            class="xs-ov"
            :style="{ left: -cv.ov!.left + 'px', width: `calc(100% + ${cv.ov!.left + cv.ov!.right}px)`, justifyContent: cv.ov!.justify }"
          >
            <span :style="{ ...cv.deco, background: cv.ov!.bg }"><CellDraft v-if="cv.draft" /><template v-else>{{ cv.text }}</template></span>
          </div>
          <span v-else-if="cv.kind !== 'icon'" :style="cv.deco"><CellDraft v-if="cv.draft" :fill="cv.draft.fill" /><template v-else>{{ cv.text }}</template></span>
        </div>
      </template>
    </td>
  </tr>
</template>
