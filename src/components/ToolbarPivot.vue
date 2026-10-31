<script setup lang="ts">
// The toolbar's PivotTable Analyze and Design commands, shown while the active cell is in a PivotTable: Options, Group
// Selection, Ungroup, Group Field and Fields, Items, & Sets (Calculated Field…, Calculated Item…); Subtotals, Grand
// Totals, Report Layout and Blank Rows with Excel's lists, the item in force checked, through the sheet's methods.
import { computed, ref, type Component } from 'vue'
import type { SheetMessages } from '../contract/messages'
import type { SheetViewApi, SheetViewState, ToolbarSource, PivotDesignChange } from '../contract/view'
import { designOf, type PivotTarget } from '../pivot-model'
import { IconDown } from './icons'
import { IconBlankRows, IconCalculated, IconGrandTotals, IconGroup, IconGroupField, IconPivotOptions, IconReportLayout, IconSubtotals, IconUngroup } from './pivot-icons'
import { useOutsidePress } from './popup'

const props = defineProps<{ group: 'pivotAnalyze' | 'pivotDesign'; sheet: SheetViewApi; state: SheetViewState | null; source: ToolbarSource; m: SheetMessages }>()
const root = ref<HTMLElement | null>(null)
const open = ref<string | null>(null)
/** What the Design lists show checked, read from the source as a list opens. */
const design = ref<ReturnType<typeof designOf> | null>(null)
useOutsidePress(root, () => (open.value = null), () => open.value !== null)
const can = computed(() => !!props.state?.canEditPivot)

async function toggle(menu: string): Promise<void> {
  open.value = open.value === menu ? null : menu
  design.value = null
  const st = props.state
  if (!open.value || props.group !== 'pivotDesign' || !st?.pivot || !props.source.pivotTable) return
  const def = await props.source.pivotTable(st.sheet, st.pivot)
  if (def) design.value = designOf({ sheet: st.sheet, pivot: st.pivot, def, cells: { labels: [], buttons: [] } } satisfies PivotTarget)
}
function run(call: () => unknown): void {
  open.value = null
  void call()
}
type Item = { id: string; label: string; on?: boolean; run: () => unknown }
const lists = computed<{ id: string; label: string; icon: Component; items: Item[] }[]>(() => {
  const x = props.m
  const d = design.value
  const set = (change: PivotDesignChange) => () => props.sheet.pivotDesign(change)
  return [
    { id: 'subtotals', label: x.pivotSubtotalsMenu, icon: IconSubtotals, items: [
      { id: 'none', label: x.pivotSubtotalsOff, on: d?.subtotals === 'none', run: set({ subtotals: 'none' }) },
      { id: 'bottom', label: x.pivotSubtotalsBottom, on: d?.subtotals === 'bottom', run: set({ subtotals: 'bottom' }) },
      { id: 'top', label: x.pivotSubtotalsTop, on: d?.subtotals === 'top', run: set({ subtotals: 'top' }) },
    ] },
    { id: 'grandTotals', label: x.pivotGrandMenu, icon: IconGrandTotals, items: [
      { id: 'off', label: x.pivotGrandOff, on: d?.grandTotals === 'off', run: set({ grandTotals: 'off' }) },
      { id: 'on', label: x.pivotGrandOn, on: d?.grandTotals === 'on', run: set({ grandTotals: 'on' }) },
      { id: 'rows', label: x.pivotGrandRowsOnly, on: d?.grandTotals === 'rows', run: set({ grandTotals: 'rows' }) },
      { id: 'columns', label: x.pivotGrandColumnsOnly, on: d?.grandTotals === 'columns', run: set({ grandTotals: 'columns' }) },
    ] },
    { id: 'layout', label: x.pivotLayoutMenu, icon: IconReportLayout, items: [
      { id: 'compact', label: x.pivotShowCompact, on: d?.layout === 'compact', run: set({ layout: 'compact' }) },
      { id: 'outline', label: x.pivotShowOutline, on: d?.layout === 'outline', run: set({ layout: 'outline' }) },
      { id: 'tabular', label: x.pivotShowTabular, on: d?.layout === 'tabular', run: set({ layout: 'tabular' }) },
      { id: 'repeat', label: x.pivotRepeatAll, on: d?.repeatLabels === true, run: set({ repeatLabels: true }) },
      { id: 'noRepeat', label: x.pivotRepeatNone, on: d?.repeatLabels === false, run: set({ repeatLabels: false }) },
    ] },
    { id: 'blankRows', label: x.pivotBlankMenu, icon: IconBlankRows, items: [
      { id: 'insert', label: x.pivotBlankInsert, on: d?.blankLine === true, run: set({ blankLine: true }) },
      { id: 'remove', label: x.pivotBlankRemove, on: d?.blankLine === false, run: set({ blankLine: false }) },
    ] },
  ]
})
</script>

<template>
  <span ref="root" class="xs-tb-pivot">
    <template v-if="group === 'pivotAnalyze'">
      <button type="button" class="xs-tool" data-tb="pivot-options" :data-tip="m.pivotOptionsTitle" :aria-label="m.pivotOptionsTitle" :disabled="!can" @click="run(() => sheet.pivotCommand('options'))"><IconPivotOptions /></button>
      <button type="button" class="xs-tool" data-tb="pivot-group" :data-tip="m.pivotGroupSelection" :aria-label="m.pivotGroupSelection" :disabled="!can" @click="run(() => sheet.groupPivot())"><IconGroup /></button>
      <button type="button" class="xs-tool" data-tb="pivot-ungroup" :data-tip="m.pivotUngroup" :aria-label="m.pivotUngroup" :disabled="!can" @click="run(() => sheet.ungroupPivot())"><IconUngroup /></button>
      <button type="button" class="xs-tool" data-tb="pivot-group-field" :data-tip="m.pivotGroupField" :aria-label="m.pivotGroupField" :disabled="!can" @click="run(() => sheet.pivotCommand('groupField'))"><IconGroupField /></button>
      <span class="xs-tb-split">
        <button type="button" class="xs-tool xs-tb-wide" data-tb="pivot-calculations" :data-tip="m.pivotFieldsItems" :aria-label="m.pivotFieldsItems" aria-haspopup="menu" :aria-expanded="open === 'calc'" :disabled="!can" @click="toggle('calc')"><IconCalculated /><IconDown /></button>
        <div v-if="open === 'calc'" class="xs-tb-pop" role="menu" :aria-label="m.pivotFieldsItems">
          <button type="button" class="xs-tb-item" role="menuitem" data-tb="pivot-calc-field" @click="run(() => sheet.pivotCommand('calculatedField'))">{{ m.pivotCalcFieldMenu }}</button>
          <button type="button" class="xs-tb-item" role="menuitem" data-tb="pivot-calc-item" @click="run(() => sheet.pivotCommand('calculatedItem'))">{{ m.pivotCalcItemMenu }}</button>
        </div>
      </span>
    </template>
    <template v-else>
      <span v-for="l in lists" :key="l.id" class="xs-tb-split">
        <button type="button" class="xs-tool xs-tb-wide" :data-tb="`pivot-${l.id}`" :data-tip="l.label" :aria-label="l.label" aria-haspopup="menu" :aria-expanded="open === l.id" :disabled="!can" @click="toggle(l.id)">
          <component :is="l.icon" /><IconDown />
        </button>
        <div v-if="open === l.id" class="xs-tb-pop" role="menu" :aria-label="l.label">
          <button v-for="it in l.items" :key="it.id" type="button" class="xs-tb-item" role="menuitemradio" :aria-checked="!!it.on" :aria-pressed="!!it.on" :data-tb="`pivot-${l.id}-${it.id}`" @click="run(it.run)">{{ it.label }}</button>
        </div>
      </span>
    </template>
  </span>
</template>
