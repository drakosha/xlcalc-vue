<script setup lang="ts">
// Excel's Workbook Statistics window: the displayed sheet's rows under Current Sheet and the workbook's under Workbook,
// a label and its value each, as Excel for Mac shows them (`statisticsRows`); OK closes it.
import { computed, useId } from 'vue'
import { statisticsRows } from '../core/statistics'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'

const ctx = useSheetContext()
const m = ctx.messages
const st = ctx.statistics
const titleId = `${useId()}-title`
const sections = computed(() => {
  const rows = statisticsRows(st.shown.value!)
  return [
    { key: 'sheet', title: m.value.statisticsSheet, rows: rows.sheet },
    { key: 'workbook', title: m.value.statisticsWorkbook, rows: rows.workbook },
  ]
})
</script>

<template>
  <DialogShell class="xs-dialog" :aria-labelledby="titleId" focus="[data-xs-dialog-ok]" data-xs-statistics @close="st.close()">
    <p :id="titleId" class="xs-modal-title">{{ m.statisticsTitle }}</p>
    <section v-for="x in sections" :key="x.key" :data-xs-stats="x.key">
      <p class="xs-stats-title">{{ x.title }}</p>
      <dl class="xs-stats-rows">
        <div v-for="[label, value] in x.rows" :key="label" class="xs-stats-row">
          <dt>{{ m[label] }}</dt>
          <dd>{{ value }}</dd>
        </div>
      </dl>
    </section>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="st.close()">{{ m.ok }}</button>
    </div>
  </DialogShell>
</template>
