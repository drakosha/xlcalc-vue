<script setup lang="ts">
// The funnel's list: a column's values with checks, a substring, the sort; under its letter, on a phone over the whole sheet.
import { useSheetContext } from './context'
import { IconClose, IconSortAsc, IconSortDesc } from './icons'
import { useOutsidePress } from './popup'

const ctx = useSheetContext()
const { filter } = ctx
const m = ctx.messages

// A press outside the list (and outside the funnels) closes it.
useOutsidePress('[data-xs-funnel],.xs-funnel', () => filter.close(), () => filter.open.value !== null)
const focusFirst = (el: unknown) => {
  if (el) (el as HTMLInputElement).focus({ preventScroll: true })
}
</script>

<template>
  <div
    v-if="filter.open.value !== null"
    data-xs-funnel
    :class="['xs-popover', ctx.v.narrow.value ? 'xs-popover-full' : '']"
    :style="ctx.v.narrow.value ? undefined : { left: filter.pos.value.left + 'px', top: filter.pos.value.top + 'px' }"
    role="dialog"
    :aria-label="filter.title.value"
    @keydown="filter.onFunnelKey"
  >
    <div class="xs-popover-head">
      <span class="xs-popover-title" :title="filter.title.value">{{ filter.title.value }}</span>
      <button
        type="button"
        :class="['xs-tool', filter.sort.value?.col === filter.open.value && filter.sort.value.dir === 'asc' ? 'xs-tool-on' : '']"
        :title="m.sortAsc"
        @click="filter.sortBy(filter.open.value!, 'asc')"
      >
        <IconSortAsc />
      </button>
      <button
        type="button"
        :class="['xs-tool', filter.sort.value?.col === filter.open.value && filter.sort.value.dir === 'desc' ? 'xs-tool-on' : '']"
        :title="m.sortDesc"
        @click="filter.sortBy(filter.open.value!, 'desc')"
      >
        <IconSortDesc />
      </button>
      <button type="button" class="xs-tool" :title="m.close" @click="filter.close()">
        <IconClose />
      </button>
    </div>
    <input :ref="ctx.v.narrow.value ? undefined : focusFirst" v-model="filter.draftText.value" class="xs-popover-input" type="text" spellcheck="false" :placeholder="m.filterText" />
    <input v-model="filter.draftSearch.value" class="xs-popover-input xs-popover-search" type="text" spellcheck="false" :placeholder="m.filterSearch" />
    <div class="xs-values" data-xs-values>
      <p v-if="filter.busy.value" class="xs-muted xs-values-note">{{ m.filterReading }}</p>
      <template v-else>
        <label class="xs-value xs-value-all">
          <input type="checkbox" :checked="filter.allChecked.value" @change="filter.toggleAll()" />
          <span class="xs-value-text">{{ m.filterAll }}</span>
          <span class="xs-muted">{{ filter.shownValues.value.length }}</span>
        </label>
        <label v-for="v in filter.shownValues.value" :key="v.text" class="xs-value">
          <input type="checkbox" :checked="filter.isChecked(v.text)" @change="filter.toggleValue(v.text)" />
          <span :class="['xs-value-text', v.text === '' ? 'xs-muted' : '']">{{ v.text === '' ? m.filterEmpty : v.text }}</span>
          <span class="xs-muted">{{ v.count }}</span>
        </label>
        <p v-if="filter.more.value" class="xs-muted xs-values-note">{{ m.filterMore(filter.more.value) }}</p>
      </template>
    </div>
    <div class="xs-popover-foot">
      <button type="button" class="xs-button xs-button-main" data-xs-apply @click="filter.applyDraft()">{{ m.filterApply }}</button>
      <button
        v-if="filter.cols.value[filter.open.value] || filter.sort.value?.col === filter.open.value"
        type="button"
        class="xs-button"
        @click="filter.clearColumn(filter.open.value!)"
      >
        {{ m.filterReset }}
      </button>
    </div>
  </div>
</template>
