<script setup lang="ts">
// Excel's Grouping dialog over a field of dates or of numbers: Auto with Starting at and Ending at, and By, the periods
// (Seconds to Years) with Number of days for days alone, or the numbers' interval. OK is xlcalc's `groupPivotField`.
import { computed } from 'vue'
import { PERIODS } from '../core/pivot-dialogs'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import '../format-dialog.css'
import '../pivot-dialogs.css'

const ctx = useSheetContext()
const m = ctx.messages
const dialogs = ctx.pivotDialogs
const d = computed(() => (dialogs.state.value?.kind === 'group' ? dialogs.state.value : null))
const daysOnly = computed(() => d.value?.draft.periods.length === 1 && d.value.draft.periods[0] === 'days')
</script>

<template>
  <DialogShell v-if="d" as="form" class="xs-dialog xs-pg" focus="[data-pg-start]" :aria-label="m.pivotGroupTitle" data-pivot-group novalidate @submit.prevent="dialogs.applyGroup()" @close="dialogs.close()">
    <p class="xs-modal-title">{{ m.pivotGroupTitle }}</p>
    <fieldset class="xs-fmt-box">
      <legend>{{ m.pivotGroupAuto }}</legend>
      <div class="xs-pg-row">
        <label class="xs-fmt-check"><input v-model="d.draft.autoStart" type="checkbox" data-pg-auto-start />{{ m.pivotGroupStart }}</label>
        <input v-model="d.draft.start" class="xs-field" type="text" spellcheck="false" :aria-label="m.pivotGroupStart" data-pg-start @input="d.draft.autoStart = false" />
      </div>
      <div class="xs-pg-row">
        <label class="xs-fmt-check"><input v-model="d.draft.autoEnd" type="checkbox" data-pg-auto-end />{{ m.pivotGroupEnd }}</label>
        <input v-model="d.draft.end" class="xs-field" type="text" spellcheck="false" :aria-label="m.pivotGroupEnd" data-pg-end @input="d.draft.autoEnd = false" />
      </div>
    </fieldset>
    <fieldset class="xs-fmt-box">
      <legend>{{ m.pivotGroupBy(d.draft.dates) }}</legend>
      <template v-if="d.draft.dates">
        <select v-model="d.draft.periods" class="xs-field xs-pv-list" multiple :size="PERIODS.length" :aria-label="m.pivotGroupBy(d.draft.dates)" data-pg-periods>
          <option v-for="p in PERIODS" :key="p" :value="p">{{ m.pivotPeriod(p) }}</option>
        </select>
        <label class="xs-pg-row">
          <span>{{ m.pivotGroupDays }}</span>
          <input v-model="d.draft.days" class="xs-field xs-po-num" type="text" inputmode="numeric" :disabled="!daysOnly" data-pg-days />
        </label>
      </template>
      <input v-else v-model="d.draft.by" class="xs-field" type="text" inputmode="decimal" :aria-label="m.pivotGroupBy(d.draft.dates)" data-pg-by />
    </fieldset>
    <p v-if="d.draft.error" class="xs-fmt-error" role="alert" data-pg-error>{{ d.draft.error }}</p>
    <div class="xs-dialog-foot">
      <button type="submit" class="xs-button xs-button-main" :disabled="d.draft.dates && !d.draft.periods.length" data-xs-dialog-ok>{{ m.ok }}</button>
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="dialogs.close()">{{ m.cancel }}</button>
    </div>
  </DialogShell>
</template>
