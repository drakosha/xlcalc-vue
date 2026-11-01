<script setup lang="ts">
// Excel's PivotTable Options over what the workbook models of it: the name; Layout & Format (merged labels, the compact
// indent, the Filters area's order and wrapping, the texts for errors and empty cells, autofit, preserved formatting);
// Totals & Filters (grand totals, filtered page items, asterisks, several filters per field, custom lists); Display; Data
// (Enable show details); Alt Text. OK is one `editPivotTable`, and a new name `renamePivotTable`.
import { computed } from 'vue'
import type { PivotOptionsTab } from '../contract/messages'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { IconClose } from './icons'
import '../format-dialog.css'
import '../pivot-dialogs.css'

const TABS: readonly PivotOptionsTab[] = ['layout', 'totals', 'display', 'data', 'altText']
const ctx = useSheetContext()
const m = ctx.messages
const dialogs = ctx.pivotDialogs
const d = computed(() => (dialogs.state.value?.kind === 'options' ? dialogs.state.value : null))
const valid = computed(() => !!d.value && dialogs.optionsValid(d.value.draft))
</script>

<template>
  <DialogShell v-if="d" as="form" class="xs-fmt xs-po" full focus="[data-po-name]" :aria-label="m.pivotOptionsTitle" data-pivot-options novalidate @submit.prevent="dialogs.applyOptions()" @close="dialogs.close()">
    <header class="xs-fmt-head">
      <span class="xs-modal-title xs-fmt-title">{{ m.pivotOptionsTitle }}</span>
      <button type="button" class="xs-tool xs-fmt-x" :title="m.cancel" :aria-label="m.cancel" @click="dialogs.close()"><IconClose /></button>
    </header>
    <div class="xs-pv-names">
      <label class="xs-fmt-field xs-po-name">
        <span>{{ m.pivotTableName }}</span>
        <input v-model="d.draft.name" class="xs-field xs-fmt-input" type="text" spellcheck="false" autocomplete="off" :aria-invalid="!d.draft.name.trim()" data-po-name />
      </label>
    </div>
    <div class="xs-fmt-tabs" role="tablist">
      <button v-for="t in TABS" :key="t" type="button" class="xs-fmt-tab" role="tab" :aria-selected="d.tab === t" :data-po-tab="t" @click="d.tab = t">{{ m.pivotOptionsTab(t) }}</button>
    </div>
    <div class="xs-fmt-body" role="tabpanel" :aria-label="m.pivotOptionsTab(d.tab)">
      <div v-show="d.tab === 'layout'" class="xs-fmt-stack">
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotLayoutGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.mergeLabels" type="checkbox" data-po="mergeLabels" />{{ m.pivotMerge }}</label>
          <label class="xs-po-row">
            <span>{{ m.pivotIndent }}</span>
            <input v-model="d.draft.indent" class="xs-field xs-po-num" type="text" inputmode="numeric" :aria-invalid="!/^\d+$/.test(d.draft.indent.trim())" data-po="indent" />
            <span>{{ m.pivotIndentUnit }}</span>
          </label>
          <label class="xs-po-row">
            <span>{{ m.pivotPageOrder }}</span>
            <select v-model="d.draft.pagesAcrossFirst" class="xs-field xs-fmt-select" data-po="pagesAcrossFirst">
              <option :value="false">{{ m.pivotDownOver }}</option>
              <option :value="true">{{ m.pivotOverDown }}</option>
            </select>
          </label>
          <label class="xs-po-row">
            <span>{{ d.draft.pagesAcrossFirst ? m.pivotPagesPerRow : m.pivotPagesPerColumn }}</span>
            <input v-model="d.draft.pagesWrap" class="xs-field xs-po-num" type="text" inputmode="numeric" data-po="pagesWrap" />
          </label>
        </fieldset>
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotFormatGroup }}</legend>
          <div class="xs-po-row">
            <label class="xs-fmt-check"><input v-model="d.draft.errorsOn" type="checkbox" data-po="errorsOn" />{{ m.pivotErrorShow }}</label>
            <input v-model="d.draft.errors" class="xs-field xs-po-text" type="text" :disabled="!d.draft.errorsOn" :aria-label="m.pivotErrorShow" data-po="errors" />
          </div>
          <div class="xs-po-row">
            <label class="xs-fmt-check"><input v-model="d.draft.emptyOn" type="checkbox" data-po="emptyOn" />{{ m.pivotEmptyShow }}</label>
            <input v-model="d.draft.empty" class="xs-field xs-po-text" type="text" :disabled="!d.draft.emptyOn" :aria-label="m.pivotEmptyShow" data-po="empty" />
          </div>
          <label class="xs-fmt-check"><input v-model="d.draft.autofitColumns" type="checkbox" data-po="autofitColumns" />{{ m.pivotAutofit }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.preserveFormatting" type="checkbox" data-po="preserveFormatting" />{{ m.pivotPreserve }}</label>
        </fieldset>
      </div>
      <div v-show="d.tab === 'totals'" class="xs-fmt-stack">
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotGrandTotalsGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.grandRows" type="checkbox" data-po="grandRows" />{{ m.pivotGrandRows }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.grandColumns" type="checkbox" data-po="grandColumns" />{{ m.pivotGrandColumns }}</label>
        </fieldset>
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotFiltersGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.subtotalHiddenItems" type="checkbox" data-po="subtotalHiddenItems" />{{ m.pivotSubtotalHidden }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.multipleFieldFilters" type="checkbox" data-po="multipleFieldFilters" />{{ m.pivotMultipleFilters }}</label>
        </fieldset>
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotSortingGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.customListSort" type="checkbox" data-po="customListSort" />{{ m.pivotCustomLists }}</label>
        </fieldset>
      </div>
      <div v-show="d.tab === 'display'" class="xs-fmt-stack">
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotDisplayGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.showDrill" type="checkbox" data-po="showDrill" />{{ m.pivotShowDrill }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.fieldHeaders" type="checkbox" data-po="fieldHeaders" />{{ m.pivotFieldCaptions }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.classicLayout" type="checkbox" data-po="classicLayout" />{{ m.pivotClassic }}</label>
          <label class="xs-fmt-check"><input v-model="d.draft.valuesRow" type="checkbox" data-po="valuesRow" />{{ m.pivotValuesRow }}</label>
        </fieldset>
      </div>
      <div v-show="d.tab === 'data'" class="xs-fmt-stack">
        <fieldset class="xs-fmt-box">
          <legend>{{ m.pivotDataGroup }}</legend>
          <label class="xs-fmt-check"><input v-model="d.draft.enableDrill" type="checkbox" data-po="enableDrill" />{{ m.pivotEnableDrill }}</label>
        </fieldset>
      </div>
      <div v-show="d.tab === 'altText'" class="xs-fmt-stack">
        <label class="xs-fmt-field">
          <span>{{ m.pivotAltTitle }}</span>
          <input v-model="d.draft.altTitle" class="xs-field xs-fmt-input" type="text" data-po="altTitle" />
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.pivotAltDescription }}</span>
          <textarea v-model="d.draft.altDescription" class="xs-field xs-po-area" rows="5" data-po="altDescription" />
        </label>
      </div>
    </div>
    <footer class="xs-fmt-foot">
      <button type="submit" class="xs-button xs-button-main xs-fmt-btn" :disabled="!valid" data-po-ok data-xs-dialog-ok>{{ m.ok }}</button>
      <button type="button" class="xs-button xs-fmt-btn" data-xs-dialog-cancel @click="dialogs.close()">{{ m.cancel }}</button>
    </footer>
  </DialogShell>
</template>
