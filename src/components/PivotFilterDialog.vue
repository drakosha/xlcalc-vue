<script setup lang="ts">
// The filter menu's dialogs: Label Filter, Value Filter and Date Filter (a comparison and its one or two boxes, the value
// field a value filter tests, Whole Days), Top 10 Filter (Top or Bottom, how many, Items, Percent or Sum, by which value
// field), and More Sort Options' Sort (Manual, Ascending or Descending by the labels or a value field). OK is one `editPivotTable`.
import { computed } from 'vue'
import { COMPARES, fieldName, isBetween, valueName } from '../pivot-model'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import '../format-dialog.css'
import '../pivot-dialogs.css'

const ctx = useSheetContext()
const m = ctx.messages
const f = ctx.pivotFilter
const d = f.dialog
const name = computed(() => (d.value ? fieldName(d.value.t, d.value.field) : ''))
const values = computed(() => (d.value?.t.def.table.values ?? []).map((_, i) => ({ i, name: valueName(d.value!.t, i) })))
const title = computed(() => {
  const x = d.value
  if (!x) return ''
  if (x.kind === 'top10') return m.value.pivotTop10Title(name.value)
  if (x.kind === 'sort') return m.value.pivotSortTitle(name.value)
  return x.filter === 'label' ? m.value.pivotLabelFilterTitle(name.value) : x.filter === 'date' ? m.value.pivotDateFilterTitle(name.value) : m.value.pivotValueFilterTitle(name.value)
})
const apply = () => {
  const x = d.value
  if (x?.kind === 'compare') void f.applyCompare()
  else if (x?.kind === 'top10') void f.applyTop10()
  else if (x?.kind === 'sort') void f.applySort()
}
/** The Sort dialog's Summary: what OK will do, as Excel describes it. */
const summary = computed(() => {
  const x = d.value
  if (x?.kind !== 'sort' || x.draft.mode === 'manual') return ''
  const by = x.draft.mode === 'ascending' ? x.draft.ascBy : x.draft.descBy
  return m.value.pivotSortSays(by < 0 ? name.value : valueName(x.t, by), x.draft.mode === 'ascending')
})
</script>

<template>
  <DialogShell v-if="d" as="form" class="xs-dialog xs-pfd-filter" focus="[data-pf-first]" :aria-label="title" data-pivot-filter-dialog novalidate @submit.prevent="apply" @close="f.closeDialog()">
    <p class="xs-modal-title">{{ title }}</p>
    <template v-if="d.kind === 'compare'">
      <p class="xs-pf-lead">{{ d.filter === 'label' ? m.pivotShowLabel : d.filter === 'date' ? m.pivotShowDate : m.pivotShowValue }}</p>
      <div class="xs-pf-row">
        <select v-if="d.filter === 'value'" v-model.number="d.draft.value" class="xs-field" :aria-label="m.pivotBy" data-pf-value>
          <option v-for="v in values" :key="v.i" :value="v.i">{{ v.name }}</option>
        </select>
        <select v-model="d.draft.compare" class="xs-field" data-pf-compare>
          <option v-for="c in COMPARES[d.filter]" :key="c" :value="c">{{ m.pivotCompare(c, d.filter) }}</option>
        </select>
        <input v-model="d.draft.first" class="xs-field" type="text" spellcheck="false" autocomplete="off" data-pf-first />
        <template v-if="isBetween(d.draft.compare)">
          <span>{{ m.pivotAnd }}</span>
          <input v-model="d.draft.second" class="xs-field" type="text" spellcheck="false" autocomplete="off" data-pf-second />
        </template>
      </div>
      <template v-if="d.filter === 'label'">
        <p class="xs-pf-hint">{{ m.pivotWildOne }}</p>
        <p class="xs-pf-hint">{{ m.pivotWildMany }}</p>
      </template>
      <label v-if="d.filter === 'date'" class="xs-fmt-check"><input v-model="d.draft.wholeDays" type="checkbox" data-pf-whole-days />{{ m.pivotWholeDays }}</label>
    </template>
    <template v-else-if="d.kind === 'top10'">
      <p class="xs-pf-lead">{{ m.pivotShow }}</p>
      <div class="xs-pf-row">
        <select v-model="d.draft.top" class="xs-field" data-pf-first data-pf-top>
          <option :value="true">{{ m.pivotTop }}</option>
          <option :value="false">{{ m.pivotBottom }}</option>
        </select>
        <input v-model="d.draft.count" class="xs-field xs-po-num" type="text" inputmode="decimal" data-pf-count />
        <select v-model="d.draft.type" class="xs-field" data-pf-type>
          <option value="count">{{ m.pivotTopItems }}</option>
          <option value="percent">{{ m.pivotTopPercent }}</option>
          <option value="sum">{{ m.pivotTopSum }}</option>
        </select>
        <span>{{ m.pivotBy }}</span>
        <select v-model.number="d.draft.value" class="xs-field" data-pf-value>
          <option v-for="v in values" :key="v.i" :value="v.i">{{ v.name }}</option>
        </select>
      </div>
    </template>
    <template v-else>
      <fieldset class="xs-fmt-box">
        <legend>{{ m.pivotSortOptions }}</legend>
        <label class="xs-fmt-check"><input v-model="d.draft.mode" type="radio" value="manual" data-pf-first data-pf-sort="manual" />{{ m.pivotSortManual }}</label>
        <label class="xs-fmt-check"><input v-model="d.draft.mode" type="radio" value="ascending" data-pf-sort="ascending" />{{ m.pivotSortAscBy }}</label>
        <select v-model.number="d.draft.ascBy" class="xs-field xs-pf-by" :disabled="d.draft.mode !== 'ascending'" data-pf-asc-by>
          <option :value="-1">{{ name }}</option>
          <option v-for="v in values" :key="v.i" :value="v.i">{{ v.name }}</option>
        </select>
        <label class="xs-fmt-check"><input v-model="d.draft.mode" type="radio" value="descending" data-pf-sort="descending" />{{ m.pivotSortDescBy }}</label>
        <select v-model.number="d.draft.descBy" class="xs-field xs-pf-by" :disabled="d.draft.mode !== 'descending'" data-pf-desc-by>
          <option :value="-1">{{ name }}</option>
          <option v-for="v in values" :key="v.i" :value="v.i">{{ v.name }}</option>
        </select>
      </fieldset>
      <fieldset class="xs-fmt-box">
        <legend>{{ m.pivotSortSummary }}</legend>
        <p class="xs-pf-hint" data-pf-summary>{{ summary }}</p>
      </fieldset>
    </template>
    <p v-if="d.kind !== 'sort' && d.error" class="xs-fmt-error" role="alert" data-pf-error>{{ d.error }}</p>
    <div class="xs-dialog-foot">
      <button type="submit" class="xs-button xs-button-main" data-xs-dialog-ok>{{ m.ok }}</button>
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="f.closeDialog()">{{ m.cancel }}</button>
    </div>
  </DialogShell>
</template>
