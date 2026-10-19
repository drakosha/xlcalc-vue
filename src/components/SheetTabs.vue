<script setup lang="ts">
// The bottom bar, as in Excel: sheet tabs (scrolling horizontally on a phone), the filter's status with Reapply and Clear, and the selection's totals.
import { computed } from 'vue'
import { useSheetContext } from './context'
import { refInto } from './dom'
import { IconFilterOff, IconPlus, IconRetry } from './icons'

const ctx = useSheetContext()
const m = ctx.messages
/** The "+" after the tabs and the tab menu, as in Excel, with the source's `structure` on an editable sheet. */
const canAdd = computed(() => ctx.cmd.can('structure'))
const { menu } = ctx
const { press } = menu

function onContext(e: MouseEvent, i: number): void {
  if (canAdd.value && press.context(e)) void menu.openSheet(i, e.clientX, e.clientY)
}
function onDown(e: PointerEvent, i: number): void {
  if (canAdd.value) press.start(e, (x, y) => void menu.openSheet(i, x, y))
}
function onClick(i: number): void {
  if (!press.after()) void ctx.pickSheet(i)
}
const setRenameInput = refInto(menu.renameInput)
</script>

<template>
  <div v-if="ctx.showTabs.value || ctx.totals.statsText.value || ctx.filter.status.value" class="xs-bottom">
    <div v-if="ctx.showTabs.value" class="xs-tabs" role="tablist">
      <template v-for="{ sheet, i } in ctx.tabs.value" :key="sheet.name + i">
        <input
          v-if="menu.renaming.value?.index === i"
          :ref="setRenameInput"
          v-model="menu.renaming.value.draft"
          class="xs-tab xs-tab-on xs-tab-input"
          :size="Math.max(4, menu.renaming.value.draft.length + 1)"
          :aria-label="m.rename"
          spellcheck="false"
          autocomplete="off"
          @keydown.enter.prevent="menu.commitRename()"
          @keydown.esc.prevent="menu.cancelRename()"
          @blur="menu.commitRename(true)"
        />
        <button
          v-else
          type="button"
          role="tab"
          :aria-selected="i === ctx.active.value"
          :class="['xs-tab', i === ctx.active.value ? 'xs-tab-on' : '']"
          :title="m.sheetTitle(sheet.name, sheet.maxRow, sheet.maxCol)"
          @click="onClick(i)"
          @dblclick="menu.startRename(i)"
          @contextmenu="onContext($event, i)"
          @pointerdown="onDown($event, i)"
          @pointermove="press.move"
          @pointerup="press.cancel"
          @pointercancel="press.cancel"
        >
          {{ sheet.name || m.sheetUnnamed(i + 1) }}
        </button>
      </template>
      <button v-if="canAdd" type="button" class="xs-tool xs-tab-add" data-xs-add-sheet :title="m.insertSheet" :aria-label="m.insertSheet" @click="ctx.menu.addSheet()">
        <IconPlus />
      </button>
    </div>
    <span v-if="ctx.filter.status.value" class="xs-filter-status" :title="ctx.filter.status.value" data-xs-filter-status>
      <span class="xs-filter-text">{{ ctx.filter.status.value }}</span>
      <button v-if="ctx.filter.available.value" type="button" class="xs-tool" :title="m.filterReapply" data-xs-filter-reapply @click="ctx.filter.reapply()">
        <IconRetry />
      </button>
      <button type="button" class="xs-tool" :title="m.filterClear" data-xs-filter-clear @click="ctx.filter.clear()">
        <IconFilterOff />
      </button>
    </span>
    <span v-if="ctx.totals.statsText.value" class="xs-stats" :title="ctx.totals.statsText.value">{{ ctx.totals.statsText.value }}</span>
  </div>
</template>
