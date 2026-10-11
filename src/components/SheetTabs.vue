<script setup lang="ts">
// The bottom line, as in Excel: sheet tabs (scrolling sideways on a phone), the filter's state and the selection's totals.
import { computed } from 'vue'
import { useSheetContext } from './context'
import { IconFilterOff, IconPlus } from './icons'

const ctx = useSheetContext()
const m = ctx.messages
/** "+" after the tabs, as Excel's, with the source's structure on an editable sheet. */
const canAdd = computed(() => ctx.cmd.can('structure'))
</script>

<template>
  <div v-if="ctx.showTabs.value || ctx.totals.statsText.value || ctx.filter.status.value" class="xs-bottom">
    <div v-if="ctx.showTabs.value" class="xs-tabs" role="tablist">
      <button
        v-for="{ sheet, i } in ctx.tabs.value"
        :key="sheet.name + i"
        type="button"
        role="tab"
        :aria-selected="i === ctx.active.value"
        :class="['xs-tab', i === ctx.active.value ? 'xs-tab-on' : '']"
        :title="m.sheetTitle(sheet.name, sheet.maxRow, sheet.maxCol)"
        @click="ctx.pickSheet(i)"
      >
        {{ sheet.name || m.sheetUnnamed(i + 1) }}
      </button>
      <button v-if="canAdd" type="button" class="xs-tool xs-tab-add" data-xs-add-sheet :title="m.insertSheet" :aria-label="m.insertSheet" @click="ctx.menu.addSheet()">
        <IconPlus />
      </button>
    </div>
    <span v-if="ctx.filter.status.value" class="xs-filter-status" :title="ctx.filter.status.value" data-xs-filter-status>
      <span class="xs-filter-text">{{ ctx.filter.status.value }}</span>
      <button type="button" class="xs-tool" :title="m.filterClear" data-xs-filter-clear @click="ctx.filter.clear()">
        <IconFilterOff />
      </button>
    </span>
    <span v-if="ctx.totals.statsText.value" class="xs-stats" :title="ctx.totals.statsText.value">{{ ctx.totals.statsText.value }}</span>
  </div>
</template>
