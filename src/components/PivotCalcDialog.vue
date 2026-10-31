<script setup lang="ts">
// Excel's Insert Calculated Field and Insert Calculated Item in "Region": Name (the existing ones listed, Modify for one of
// them), Formula, the fields (and a field's items) with Insert Field and Insert Item, Add or Modify, Delete; OK sets the
// field or item named and closes, Close closes. Each Add, Modify, Delete and OK is one operation of xlcalc's.
import { computed, ref, useId, watch } from 'vue'
import { axisFields, fieldName } from '../pivot-model'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import '../format-dialog.css'
import '../pivot-dialogs.css'

const ctx = useSheetContext()
const m = ctx.messages
const dialogs = ctx.pivotDialogs
const d = computed(() => (dialogs.state.value?.kind === 'calcField' || dialogs.state.value?.kind === 'calcItem' ? dialogs.state.value : null))
const listId = `${useId()}-names`
const formulaBox = ref<HTMLInputElement | null>(null)

/** The calculated fields of the cache, or the calculated items of the field, with their formulas as Excel's box shows them. */
const known = computed(() => {
  const x = d.value
  if (!x) return []
  const cache = x.t.def.cache
  if (x.kind === 'calcField') return cache.fields.flatMap((f) => (f.formula !== undefined ? [{ name: f.name, formula: `=${f.formula}` }] : []))
  const values = cache.fields[x.field]?.items ?? []
  return (cache.calculatedItems ?? []).filter((c) => c.field === x.field).map((c) => ({ name: String(values[c.item] ?? ''), formula: `=${c.formula.replace(new RegExp(`${fieldName(x.t, x.field)}\\[([^\\]]*)\\]`, 'g'), '$1')}` }))
})
const existing = computed(() => known.value.find((k) => k.name.toLowerCase() === d.value?.draft.name.trim().toLowerCase()))
// a name of the list picked shows its formula, as Excel's Name box does
watch(existing, (k) => k && d.value && (d.value.draft.formula = k.formula))

/** The fields listed: the source's for a calculated field, the PivotTable's on its rows and columns for an item. */
const fields = computed(() => {
  const x = d.value
  if (!x) return []
  const list = x.kind === 'calcField' ? x.t.def.cache.fields.map((_, i) => i) : [...axisFields(x.t, 'rows'), ...axisFields(x.t, 'columns')]
  return list.map((i) => ({ i, name: fieldName(x.t, i) }))
})
const items = computed(() => (d.value?.kind === 'calcItem' ? (d.value.t.cells.labels[d.value.draft.pick] ?? []) : []))
/** A name in a formula as Excel's dialog writes it: in quotes unless it is a plain name. */
const quoted = (name: string) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`)
function insert(text: string): void {
  const x = d.value
  if (!x) return
  const el = formulaBox.value
  const at = el?.selectionStart ?? x.draft.formula.length
  const end = el?.selectionEnd ?? at
  x.draft.formula = x.draft.formula.slice(0, at) + text + x.draft.formula.slice(end)
}
const insertField = () => d.value && insert(quoted(fieldName(d.value.t, d.value.draft.pick)))
const insertItem = () => d.value && items.value[d.value.draft.item] !== undefined && insert(quoted(items.value[d.value.draft.item]!))
</script>

<template>
  <DialogShell
    v-if="d"
    as="form"
    class="xs-dialog xs-pc"
    focus="[data-pc-name]"
    :aria-label="d.kind === 'calcField' ? m.pivotCalcFieldTitle : m.pivotCalcItemTitle(fieldName(d.t, d.field))"
    data-pivot-calc
    novalidate
    @submit.prevent="dialogs.applyCalc(true)"
    @close="dialogs.close()"
  >
    <p class="xs-modal-title">{{ d.kind === 'calcField' ? m.pivotCalcFieldTitle : m.pivotCalcItemTitle(fieldName(d.t, d.field)) }}</p>
    <div class="xs-pc-grid">
      <label class="xs-fmt-field">
        <span>{{ m.pivotName }}</span>
        <input v-model="d.draft.name" class="xs-field" type="text" spellcheck="false" autocomplete="off" :list="listId" data-pc-name />
        <datalist :id="listId">
          <option v-for="k in known" :key="k.name" :value="k.name" />
        </datalist>
      </label>
      <button type="button" class="xs-button" :disabled="!d.draft.name.trim()" data-pc-add @click="dialogs.applyCalc(false)">{{ existing ? m.pivotModify : m.pivotAdd }}</button>
      <label class="xs-fmt-field">
        <span>{{ m.pivotFormula }}</span>
        <input ref="formulaBox" v-model="d.draft.formula" class="xs-field" type="text" spellcheck="false" autocomplete="off" data-pc-formula />
      </label>
      <button type="button" class="xs-button" :disabled="!existing" data-pc-delete @click="dialogs.deleteCalc()">{{ m.pivotDelete }}</button>
    </div>
    <div class="xs-pc-lists">
      <label class="xs-fmt-field">
        <span>{{ m.pivotFields }}</span>
        <select v-model.number="d.draft.pick" class="xs-field xs-pv-list" size="8" data-pc-fields @dblclick="insertField">
          <option v-for="f in fields" :key="f.i" :value="f.i">{{ f.name }}</option>
        </select>
        <button type="button" class="xs-button" data-pc-insert-field @click="insertField">{{ m.pivotInsertField }}</button>
      </label>
      <label v-if="d.kind === 'calcItem'" class="xs-fmt-field">
        <span>{{ m.pivotItems }}</span>
        <select v-model.number="d.draft.item" class="xs-field xs-pv-list" size="8" data-pc-items @dblclick="insertItem">
          <option v-for="(label, i) in items" :key="i" :value="i">{{ label }}</option>
        </select>
        <button type="button" class="xs-button" data-pc-insert-item @click="insertItem">{{ m.pivotInsertItem }}</button>
      </label>
    </div>
    <div class="xs-dialog-foot">
      <button type="submit" class="xs-button xs-button-main" data-xs-dialog-ok>{{ m.ok }}</button>
      <button type="button" class="xs-button" data-pc-close @click="dialogs.close()">{{ m.close }}</button>
    </div>
  </DialogShell>
</template>
