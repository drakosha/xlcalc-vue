<script setup lang="ts">
// The bottom line, as in Excel: sheet tabs (scrolling sideways on a phone) and the selection's totals.
import { useSheetContext } from './context'

const ctx = useSheetContext()
const m = ctx.messages
</script>

<template>
  <div v-if="ctx.tabs.value.length > 1 || ctx.totals.statsText.value" class="xs-bottom">
    <div v-if="ctx.tabs.value.length > 1" class="xs-tabs" role="tablist">
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
    </div>
    <span v-if="ctx.totals.statsText.value" class="xs-stats" :title="ctx.totals.statsText.value">{{ ctx.totals.statsText.value }}</span>
  </div>
</template>
