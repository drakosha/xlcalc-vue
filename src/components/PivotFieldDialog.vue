<script setup lang="ts">
// Excel's Field Settings of a field on the rows, the columns or the Filters area: Custom Name; Subtotals & Filters
// (Automatic, None or Custom functions, Include new items in manual filter); Layout & Print (outline form with its compact
// and top subtotals, or tabular form, repeated labels, a blank line, items with no data, a page break). OK is one `editPivotTable`.
import { computed } from 'vue'
import type { PivotFunction } from 'xlcalc/contract'
import { PIVOT_FUNCTIONS } from '../pivot-model'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { IconClose } from './icons'
import '../format-dialog.css'
import '../pivot-dialogs.css'

const ctx = useSheetContext()
const m = ctx.messages
const dialogs = ctx.pivotDialogs
const d = computed(() => (dialogs.state.value?.kind === 'field' ? dialogs.state.value : null))
const custom = computed({
  get: () => d.value?.draft.custom ?? [],
  set: (v: PivotFunction[]) => d.value && (d.value.draft.custom = v),
})
/** Custom without a function chosen leaves nothing to show, so OK waits for one. */
const blocked = computed(() => d.value?.draft.subtotals === 'custom' && !d.value.draft.custom.length)
</script>

<template>
  <DialogShell v-if="d" as="form" class="xs-fmt xs-pfd" full focus="[data-pfd-name]" :aria-label="m.pivotFieldTitle" data-pivot-field novalidate @submit.prevent="!blocked && dialogs.applyField()" @close="dialogs.close()">
    <header class="xs-fmt-head">
      <span class="xs-modal-title xs-fmt-title">{{ m.pivotFieldTitle }}</span>
      <button type="button" class="xs-tool xs-fmt-x" :title="m.cancel" :aria-label="m.cancel" @click="dialogs.close()"><IconClose /></button>
    </header>
    <div class="xs-pv-names">
      <p class="xs-pv-source">{{ m.pivotSourceName }} <span data-pfd-source>{{ d.t.def.cache.fields[d.field]?.name }}</span></p>
      <label class="xs-fmt-field">
        <span>{{ m.pivotCustomName }}</span>
        <input v-model="d.draft.name" class="xs-field xs-fmt-input" type="text" spellcheck="false" autocomplete="off" data-pfd-name />
      </label>
    </div>
    <div class="xs-fmt-tabs" role="tablist">
      <button type="button" class="xs-fmt-tab" role="tab" :aria-selected="d.tab === 'subtotals'" data-pfd-tab="subtotals" @click="d.tab = 'subtotals'">{{ m.pivotSubtotalsTab }}</button>
      <button type="button" class="xs-fmt-tab" role="tab" :aria-selected="d.tab === 'layout'" data-pfd-tab="layout" @click="d.tab = 'layout'">{{ m.pivotLayoutTab }}</button>
    </div>
    <div class="xs-fmt-body" role="tabpanel">
      <div v-show="d.tab === 'subtotals'" class="xs-fmt-stack">
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotSubtotals }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.subtotals" type="radio" value="auto" data-pfd-subtotals="auto" />{{ m.pivotSubtotalsAuto }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.subtotals" type="radio" value="none" data-pfd-subtotals="none" />{{ m.pivotSubtotalsNone }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.subtotals" type="radio" value="custom" data-pfd-subtotals="custom" />{{ m.pivotSubtotalsCustom }}</label>
          <label class="xs-fmt-field xs-pfd-pick">
            <span>{{ m.pivotSubtotalsPick }}</span>
            <select v-model="custom" class="xs-field xs-pv-list" multiple :size="PIVOT_FUNCTIONS.length" :disabled="d.draft.subtotals !== 'custom'" data-pfd-custom>
              <option v-for="fn in PIVOT_FUNCTIONS" :key="fn" :value="fn">{{ m.pivotFunction(fn) }}</option>
            </select>
          </label>
        </fieldset>
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotFilterGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.includeNewItems" type="checkbox" data-pfd-include-new />{{ m.pivotIncludeNew }}</label>
        </fieldset>
      </div>
      <div v-show="d.tab === 'layout'" class="xs-fmt-stack">
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotLayoutGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.form" type="radio" value="outline" data-pfd-form="outline" />{{ m.pivotOutlineForm }}</label>
          <div class="xs-pfd-inner">
            <label class="xs-fmt-check"><input v-model="d.draft.compact" type="checkbox" :disabled="d.draft.form !== 'outline'" data-pfd-compact />{{ m.pivotCompactForm }}</label>
            <label class="xs-fmt-check"><input v-model="d.draft.subtotalTop" type="checkbox" :disabled="d.draft.form !== 'outline'" data-pfd-top />{{ m.pivotSubtotalTop }}</label>
          </div>
          <label class="xs-fmt-check"><input v-model="d.draft.form" type="radio" value="tabular" data-pfd-form="tabular" />{{ m.pivotTabularForm }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.repeatLabels" type="checkbox" data-pfd-repeat />{{ m.pivotRepeatLabels }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.blankLine" type="checkbox" data-pfd-blank />{{ m.pivotBlankLine }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.showEmpty" type="checkbox" data-pfd-empty />{{ m.pivotShowEmpty }}</label>
        </fieldset>
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotPrintGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.pageBreak" type="checkbox" data-pfd-page-break />{{ m.pivotPageBreak }}</label>
        </fieldset>
      </div>
    </div>
    <footer class="xs-fmt-foot">
      <button type="submit" class="xs-button xs-button-main xs-fmt-btn" :disabled="blocked" data-pfd-ok data-xs-dialog-ok>{{ m.ok }}</button>
      <button type="button" class="xs-button xs-fmt-btn" data-xs-dialog-cancel @click="dialogs.close()">{{ m.cancel }}</button>
    </footer>
  </DialogShell>
</template>
