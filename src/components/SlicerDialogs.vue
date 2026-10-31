<script setup lang="ts">
// Excel's dialogs of slicers and timelines: Insert Slicers and Insert Timelines (a box for each field of the PivotTable's
// cache, of dates for a timeline), Slicer Settings (the names, Header, Item Sorting and Filtering) and Report Connections
// (the PivotTables of the same pivot cache, by name and sheet). OK applies, as one operation each; Esc or Cancel closes.
import { computed, useId } from 'vue'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import '../controls.css'
import '../slicers.css'

const ctx = useSheetContext()
const { slicers } = ctx
const m = ctx.messages
const titleId = `${useId()}-title`
const sortName = `${useId()}-sort`

const insert = computed(() => (slicers.dialog.value?.kind === 'insert' ? slicers.dialog.value : null))
const settings = computed(() => (slicers.dialog.value?.kind === 'settings' ? slicers.dialog.value : null))
const connections = computed(() => (slicers.dialog.value?.kind === 'connections' ? slicers.dialog.value : null))
const title = computed(() => (insert.value ? (insert.value.timeline ? m.value.insertTimelinesTitle : m.value.insertSlicersTitle) : settings.value ? m.value.slicerSettingsTitle : m.value.reportConnectionsTitle(connections.value?.object.name ?? '')))
function ok(): void {
  if (insert.value) void slicers.insert()
  else if (settings.value) void slicers.applySettings()
  else void slicers.applyConnections()
}
</script>

<template>
  <DialogShell class="xs-dialog xs-slicer-dialog" :aria-labelledby="titleId" focus="[data-xs-dialog-ok]" data-xs-slicer-dialog @close="slicers.close()">
    <p :id="titleId" class="xs-modal-title">{{ title }}</p>
    <div v-if="insert" class="xs-dialog-list xs-slicer-fields" data-xs-slicer-fields>
      <label v-for="f in insert.fields" :key="f.name" class="xs-dialog-choice"><input v-model="f.checked" type="checkbox" :data-xs-field="f.name" />{{ f.name }}</label>
    </div>
    <template v-else-if="settings">
      <p class="xs-slicer-dialog-line">{{ m.slicerSourceName }} {{ settings.field }}</p>
      <p class="xs-slicer-dialog-line">{{ m.slicerFormulaName }} {{ settings.cache }}</p>
      <label class="xs-dialog-field">
        <span>{{ m.slicerName }}</span>
        <input class="xs-field" type="text" :value="settings.slicer" disabled data-xs-slicer-name />
      </label>
      <fieldset class="xs-dialog-choices">
        <legend>{{ m.slicerHeader }}</legend>
        <label class="xs-dialog-choice"><input v-model="settings.showCaption" type="checkbox" data-xs-slicer-show-caption />{{ m.slicerDisplayHeader }}</label>
        <label class="xs-dialog-field">
          <span>{{ m.slicerCaption }}</span>
          <input v-model="settings.caption" class="xs-field" type="text" autocomplete="off" :disabled="!settings.showCaption" data-xs-slicer-caption @keydown.enter.prevent="ok" />
        </label>
      </fieldset>
      <fieldset class="xs-dialog-choices">
        <legend>{{ m.slicerSorting }}</legend>
        <label class="xs-dialog-choice"><input v-model="settings.descending" type="radio" :name="sortName" :value="false" data-xs-slicer-ascending />{{ m.slicerAscending }}</label>
        <label class="xs-dialog-choice"><input v-model="settings.descending" type="radio" :name="sortName" :value="true" data-xs-slicer-descending />{{ m.slicerDescending }}</label>
        <label class="xs-dialog-choice"><input v-model="settings.customListSort" type="checkbox" data-xs-slicer-custom-lists />{{ m.slicerCustomLists }}</label>
        <label class="xs-dialog-choice"><input v-model="settings.hide" type="checkbox" data-xs-slicer-hide />{{ m.slicerHideNoData }}</label>
        <label class="xs-dialog-choice"><input v-model="settings.mark" type="checkbox" :disabled="settings.hide" data-xs-slicer-mark />{{ m.slicerMarkNoData }}</label>
        <label class="xs-dialog-choice"><input v-model="settings.last" type="checkbox" :disabled="settings.hide || !settings.mark" data-xs-slicer-last />{{ m.slicerNoDataLast }}</label>
        <label class="xs-dialog-choice"><input v-model="settings.showMissing" type="checkbox" data-xs-slicer-show-missing />{{ m.slicerShowMissing }}</label>
      </fieldset>
    </template>
    <template v-else-if="connections">
      <p class="xs-slicer-dialog-line">{{ m.reportConnectionsPrompt }}</p>
      <table class="xs-slicer-connections" data-xs-connections>
        <thead>
          <tr>
            <th></th>
            <th>{{ m.reportConnectionsName }}</th>
            <th>{{ m.reportConnectionsSheet }}</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="p in connections.pivots" :key="`${p.sheet}!${p.name}`">
            <td><input v-model="p.checked" type="checkbox" :aria-label="`${p.name} ${p.sheet}`" :data-xs-connection="`${p.sheet}!${p.name}`" /></td>
            <td>{{ p.name }}</td>
            <td>{{ p.sheet }}</td>
          </tr>
        </tbody>
      </table>
    </template>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button xs-button-main" :disabled="!!insert && !insert.fields.some((f) => f.checked)" data-xs-dialog-ok @click="ok">{{ m.ok }}</button>
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="slicers.close()">{{ m.cancel }}</button>
    </div>
  </DialogShell>
</template>

