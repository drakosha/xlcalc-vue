<script setup lang="ts">
// Excel's filter menu of a PivotTable's field: Select field (compact form), Sort A to Z and Z to A, More Sort Options…,
// Clear Filter From, Label Filters or Date Filters and Value Filters with their submenus, the search box and the items'
// list with (Select All); a Filters area's field lists (All) and its items, or with Select Multiple Items boxes to check.
import { computed, nextTick, ref, watch } from 'vue'
import type { PivotCompare } from '../contract/messages'
import { COMPARES, DATE_MENU, DATE_PERIODS, fieldName, filtered, type FilterKind } from '../pivot-model'
import { useSheetContext } from './context'
import { IconSortAsc, IconSortDesc } from './icons'
import { useOutsidePress } from './popup'
import '../controls.css'
import '../filter.css'
import '../menu.css'
import '../pivot-dialogs.css'

const ctx = useSheetContext()
const m = ctx.messages
const f = ctx.pivotFilter
const x = f.menu
const box = ref<HTMLElement | null>(null)
const pos = ref({ left: '0px', top: '0px' })
/** The open submenu: a kind of filter, or All Dates in the Period within Date Filters. */
const sub = ref<FilterKind | null>(null)
const allDates = ref(false)

watch(
  () => x.value?.at,
  async (at) => {
    sub.value = null
    allDates.value = false
    if (!at) return
    pos.value = { left: `${at.left}px`, top: `${at.top}px` }
    await nextTick()
    const r = box.value?.getBoundingClientRect()
    if (!r) return
    pos.value = { left: `${Math.max(4, Math.min(at.left, window.innerWidth - r.width - 4))}px`, top: `${Math.max(4, Math.min(at.top, window.innerHeight - r.height - 4))}px` }
    box.value?.querySelector<HTMLElement>('[data-pfm-first]')?.focus({ preventScroll: true })
  },
)
useOutsidePress(box, () => f.closeMenu(), () => !!x.value)

const t = computed(() => x.value?.t)
const name = computed(() => (x.value ? fieldName(x.value.t, x.value.field) : ''))
const labelKind = computed<FilterKind>(() => (x.value ? f.labelKind(x.value.t, x.value.field) : 'label'))
const hasValues = computed(() => !!t.value?.def.table.values?.length)
const list = computed(() => f.listed())
const allChecked = computed(() => list.value.every((y) => !x.value?.hidden[y.i]))
function toggleAll(): void {
  const on = !allChecked.value
  for (const y of list.value) x.value!.hidden[y.i] = !on
}
const compares = (kind: FilterKind): readonly PivotCompare[] => (kind === 'date' ? DATE_MENU : COMPARES[kind])
function onKey(e: KeyboardEvent): void {
  if (e.key !== 'Escape') return
  e.preventDefault()
  e.stopPropagation()
  f.closeMenu()
  ctx.focusGrid()
}
</script>

<template>
  <div v-if="x && t" ref="box" class="xs-popover xs-pfm" role="dialog" :aria-label="name" :style="pos" data-pivot-filter-menu @keydown="onKey">
    <label v-if="x.fields.length > 1" class="xs-pfm-field">
      <span>{{ m.pivotSelectField }}</span>
      <select class="xs-field" :value="x.field" data-pfm-field @change="f.pick(Number(($event.target as HTMLSelectElement).value))">
        <option v-for="i in x.fields" :key="i" :value="i">{{ fieldName(t, i) }}</option>
      </select>
    </label>
    <template v-if="x.axis !== 'pages'">
      <div class="xs-pfm-items" role="menu">
        <button type="button" role="menuitem" class="xs-menu-item xs-pfm-item" data-pfm-first data-pfm="sortAsc" @click="f.sort('ascending')"><IconSortAsc />{{ labelKind === 'date' ? m.pivotSortOldest : m.pivotSortAsc }}</button>
        <button type="button" role="menuitem" class="xs-menu-item xs-pfm-item" data-pfm="sortDesc" @click="f.sort('descending')"><IconSortDesc />{{ labelKind === 'date' ? m.pivotSortNewest : m.pivotSortDesc }}</button>
        <button type="button" role="menuitem" class="xs-menu-item xs-pfm-item" data-pfm="moreSort" @click="f.openSort()">{{ m.pivotMoreSort }}</button>
        <div class="xs-menu-sep" role="separator"></div>
        <button type="button" role="menuitem" class="xs-menu-item xs-pfm-item" :disabled="!filtered(t, x.field)" data-pfm="clear" @click="f.clear()">{{ m.pivotClearFrom(name) }}</button>
        <div v-for="kind in [labelKind, 'value'] as FilterKind[]" :key="kind" class="xs-pfm-sub" @mouseenter="sub = kind">
          <button
            type="button"
            role="menuitem"
            class="xs-menu-item xs-pfm-item xs-pfm-more"
            aria-haspopup="menu"
            :aria-expanded="sub === kind"
            :disabled="kind === 'value' && !hasValues"
            :data-pfm="kind"
            @click="sub = sub === kind ? null : kind"
          >
            {{ kind === 'label' ? m.pivotLabelFilters : kind === 'date' ? m.pivotDateFilters : m.pivotValueFilters }}
          </button>
          <div v-if="sub === kind" class="xs-menu xs-pfm-flyout" role="menu" :data-pfm-sub="kind">
            <button type="button" role="menuitem" class="xs-menu-item" :disabled="!t.def.table.filters?.some((y) => y.field === x!.field)" data-pfm-op="clear" @click="f.clear()">{{ m.pivotClearFilter }}</button>
            <div class="xs-menu-sep" role="separator"></div>
            <button v-for="c in compares(kind)" :key="c" type="button" role="menuitem" class="xs-menu-item" :data-pfm-op="c" @click="f.openCompare(kind, c)">{{ m.pivotCompareMenu(c, kind) }}</button>
            <template v-if="kind === 'value'">
              <div class="xs-menu-sep" role="separator"></div>
              <button type="button" role="menuitem" class="xs-menu-item" data-pfm-op="top10" @click="f.openTop10()">{{ m.pivotTop10Menu }}</button>
            </template>
            <template v-if="kind === 'date'">
              <div class="xs-menu-sep" role="separator"></div>
              <button v-for="p in DATE_PERIODS" :key="p" type="button" role="menuitem" class="xs-menu-item" :data-pfm-op="p" @click="f.period(p)">{{ m.pivotDatePeriod(p) }}</button>
              <div class="xs-pfm-sub" @mouseenter="allDates = true" @mouseleave="allDates = false">
                <button type="button" role="menuitem" class="xs-menu-item xs-pfm-more" aria-haspopup="menu" :aria-expanded="allDates" data-pfm-op="allDates" @click="allDates = !allDates">{{ m.pivotAllDates }}</button>
                <div v-if="allDates" class="xs-menu xs-pfm-flyout" role="menu" data-pfm-sub="allDates">
                  <button v-for="q in 4" :key="q" type="button" role="menuitem" class="xs-menu-item" :data-pfm-op="`Q${q}`" @click="f.period(`Q${q}` as 'Q1')">{{ m.pivotQuarter(q) }}</button>
                  <div class="xs-menu-sep" role="separator"></div>
                  <button v-for="mo in 12" :key="mo" type="button" role="menuitem" class="xs-menu-item" :data-pfm-op="`M${mo}`" @click="f.period(`M${mo}` as 'M1')">{{ m.pivotMonth(mo) }}</button>
                </div>
              </div>
              <button type="button" role="menuitem" class="xs-menu-item" data-pfm-op="custom" @click="f.openCompare('date', 'equal')">{{ m.pivotCustomFilter }}</button>
            </template>
          </div>
        </div>
      </div>
    </template>
    <label v-else class="xs-fmt-check xs-pfm-multiple"><input v-model="x.multiple" type="checkbox" data-pfm-multiple />{{ m.pivotMultipleItems }}</label>
    <input v-model="x.search" class="xs-field xs-popover-input" type="text" spellcheck="false" :placeholder="m.pivotSearch" :aria-label="m.pivotSearch" :data-pfm-first="x.axis === 'pages' ? '' : undefined" data-pfm-search />
    <div class="xs-values xs-pfm-values" data-pfm-values>
      <template v-if="x.axis === 'pages' && !x.multiple">
        <label class="xs-value"><input v-model="x.page" type="radio" :value="-1" data-pfm-page="-1" /><span class="xs-value-text">{{ m.pivotAll }}</span></label>
        <label v-for="y in list" :key="y.i" class="xs-value"><input v-model="x.page" type="radio" :value="y.i" :data-pfm-page="y.i" /><span class="xs-value-text">{{ y.label }}</span></label>
      </template>
      <template v-else>
        <label class="xs-value xs-value-all"><input type="checkbox" :checked="allChecked" data-pfm-all @change="toggleAll" /><span class="xs-value-text">{{ m.pivotSelectAll }}</span></label>
        <label v-for="y in list" :key="y.i" class="xs-value">
          <input type="checkbox" :checked="!x.hidden[y.i]" :data-pfm-item="y.i" @change="x.hidden[y.i] = !($event.target as HTMLInputElement).checked" />
          <span class="xs-value-text">{{ y.label }}</span>
        </label>
      </template>
    </div>
    <div class="xs-popover-foot xs-pfm-foot">
      <button type="button" class="xs-button xs-button-main" :disabled="!f.canApply()" data-pfm-ok @click="f.applyList()">{{ m.ok }}</button>
      <button type="button" class="xs-button" data-pfm-cancel @click="f.closeMenu(), ctx.focusGrid()">{{ m.cancel }}</button>
    </div>
  </div>
</template>
