<script setup lang="ts">
// Excel's Value Field Settings: the source field's name, Custom Name, Summarize Values By (Excel's 11 functions) and Show
// Values As (its 15 calculations with their base field and base item), and Number Format (Format Cells' Number tab); OK
// sends the value field as one `editPivotTable`.
import { computed } from 'vue'
import type { PivotFunction } from 'xlcalc/contract'
import { fieldName, PIVOT_FUNCTIONS } from '../pivot-model'
import { BASE_FIELD, BASE_ITEM, SHOW_AS } from '../core/pivot-dialogs'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { IconClose } from './icons'
import '../format-dialog.css'
import '../pivot-dialogs.css'

const ctx = useSheetContext()
const m = ctx.messages
const dialogs = ctx.pivotDialogs
const d = computed(() => (dialogs.state.value?.kind === 'value' ? dialogs.state.value : null))
/** Every field of the source but the calculated ones may be a base field. */
const bases = computed(() => (d.value ? d.value.t.def.cache.fields.flatMap((f, i) => (f.formula === undefined ? [{ i, name: fieldName(d.value!.t, i) }] : [])) : []))
const items = computed(() => (d.value ? (d.value.t.cells.labels[d.value.draft.baseField] ?? []) : []))
const baseItem = computed({
  get: () => String(d.value?.draft.baseItem ?? 'previous'),
  set: (v: string) => d.value && (d.value.draft.baseItem = v === 'previous' || v === 'next' ? v : Number(v)),
})
</script>

<template>
  <DialogShell v-if="d" as="form" class="xs-fmt xs-pv" full focus="[data-pv-name]" :aria-label="m.pivotValueTitle" data-pivot-value novalidate @submit.prevent="dialogs.applyValue()" @close="dialogs.close()">
    <header class="xs-fmt-head">
      <span class="xs-modal-title xs-fmt-title">{{ m.pivotValueTitle }}</span>
      <button type="button" class="xs-tool xs-fmt-x" :title="m.cancel" :aria-label="m.cancel" @click="dialogs.close()"><IconClose /></button>
    </header>
    <div class="xs-pv-names">
      <p class="xs-pv-source">{{ m.pivotSourceName }} <span data-pv-source>{{ d.t.def.cache.fields[d.t.def.table.values![d.index]!.field]?.name }}</span></p>
      <label class="xs-fmt-field">
        <span>{{ m.pivotCustomName }}</span>
        <input v-model="d.draft.name" class="xs-field xs-fmt-input" type="text" spellcheck="false" autocomplete="off" data-pv-name />
      </label>
    </div>
    <div class="xs-fmt-tabs" role="tablist">
      <button type="button" class="xs-fmt-tab" role="tab" :aria-selected="d.tab === 'summarize'" data-pv-tab="summarize" @click="d.tab = 'summarize'">{{ m.pivotSummarizeTab }}</button>
      <button type="button" class="xs-fmt-tab" role="tab" :aria-selected="d.tab === 'showAs'" data-pv-tab="showAs" @click="d.tab = 'showAs'">{{ m.pivotShowAsTab }}</button>
    </div>
    <div class="xs-fmt-body xs-pv-body" role="tabpanel">
      <fieldset v-show="d.tab === 'summarize'" class="xs-fmt-box">
        <legend>{{ m.pivotSummarizeBy }}</legend>
        <p class="xs-pv-hint">{{ m.pivotSummarizeHint }}</p>
        <select :value="d.draft.fn" class="xs-field xs-pv-list" :size="PIVOT_FUNCTIONS.length" :aria-label="m.pivotSummarizeBy" data-pv-function @change="dialogs.pickFunction(($event.target as HTMLSelectElement).value as PivotFunction)">
          <option v-for="fn in PIVOT_FUNCTIONS" :key="fn" :value="fn">{{ m.pivotFunction(fn) }}</option>
        </select>
      </fieldset>
      <div v-show="d.tab === 'showAs'" class="xs-fmt-stack">
        <label class="xs-fmt-field">
          <span>{{ m.pivotShowValuesAs }}</span>
          <select v-model="d.draft.showAs" class="xs-field xs-fmt-select" data-pv-show-as>
            <option v-for="c in SHOW_AS" :key="c" :value="c">{{ m.pivotShowAs(c) }}</option>
          </select>
        </label>
        <div class="xs-pv-bases">
          <label class="xs-fmt-field">
            <span>{{ m.pivotBaseField }}</span>
            <select :value="d.draft.baseField" class="xs-field xs-pv-list" size="8" :disabled="!BASE_FIELD.has(d.draft.showAs)" data-pv-base-field @change="dialogs.pickBaseField(Number(($event.target as HTMLSelectElement).value))">
              <option v-for="b in bases" :key="b.i" :value="b.i">{{ b.name }}</option>
            </select>
          </label>
          <label class="xs-fmt-field">
            <span>{{ m.pivotBaseItem }}</span>
            <select v-model="baseItem" class="xs-field xs-pv-list" size="8" :disabled="!BASE_ITEM.has(d.draft.showAs)" data-pv-base-item>
              <option value="previous">{{ m.pivotPrevious }}</option>
              <option value="next">{{ m.pivotNext }}</option>
              <option v-for="(label, i) in items" :key="i" :value="String(i)">{{ label }}</option>
            </select>
          </label>
        </div>
      </div>
    </div>
    <footer class="xs-fmt-foot">
      <button type="button" class="xs-button xs-fmt-btn xs-pv-format" data-pv-format @click="dialogs.openNumberFormat(d.t, d.index)">{{ m.pivotNumberFormat }}</button>
      <button type="submit" class="xs-button xs-button-main xs-fmt-btn" data-pv-ok data-xs-dialog-ok>{{ m.ok }}</button>
      <button type="button" class="xs-button xs-fmt-btn" data-xs-dialog-cancel @click="dialogs.close()">{{ m.cancel }}</button>
    </footer>
  </DialogShell>
</template>
