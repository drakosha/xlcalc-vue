<script setup lang="ts">
// A slicer as Excel draws it, in a plain neutral look (its style is kept for Excel and not drawn): the header with its
// caption, Multi-Select and Clear Filter, then a button for each item in its columns, chosen, not chosen or without data,
// scrolled when they do not fit. A click chooses an item, Ctrl+click and Multi-Select add or take one away, Shift+click a run.
import { computed } from 'vue'
import type { SheetSlicer } from '../contract/source'
import { EMU_PX } from '../core/drawing'
import { useSheetContext } from './context'
import { IconFilterOff, IconMultiSelect } from './icons'

const props = defineProps<{ slicer: SheetSlicer }>()
const ctx = useSheetContext()
const { slicers } = ctx
const m = ctx.messages

/** Excel's button height without one of its own: 0.26 inch. */
const ROW_EMU = 241_300
const items = computed(() => slicers.shownItems(props.slicer))
const style = computed(() => ({ gridTemplateColumns: `repeat(${Math.max(1, props.slicer.columns ?? 1)}, minmax(0, 1fr))`, gridAutoRows: `${Math.round((props.slicer.rowHeight ?? ROW_EMU) / EMU_PX)}px` }))
const multi = computed(() => !!slicers.multi.value[props.slicer.name])
const filtered = computed(() => slicers.filtered(props.slicer))
const usable = computed(() => slicers.canChoose(props.slicer))
const mark = computed(() => props.slicer.noData?.mark !== false)
const ref_ = () => ({ kind: 'slicer' as const, name: props.slicer.name })

function pick(e: MouseEvent, item: number): void {
  slicers.select(ref_())
  void slicers.choose(props.slicer, item, { ctrl: e.ctrlKey || e.metaKey, shift: e.shiftKey })
}
</script>

<template>
  <div class="xs-slicer" :data-xs-slicer="slicer.name">
    <div v-if="slicer.showCaption !== false" class="xs-slicer-head">
      <span class="xs-slicer-caption">{{ slicer.caption ?? slicer.name }}</span>
      <button type="button" :class="['xs-slicer-tool', { 'xs-on': multi }]" :title="m.slicerMultiSelect" :aria-label="m.slicerMultiSelect" :aria-pressed="multi" data-xs-slicer-multi @click="slicers.toggleMulti(slicer.name)">
        <IconMultiSelect />
      </button>
      <button type="button" class="xs-slicer-tool" :title="m.slicerClearFilter" :aria-label="m.slicerClearFilter" :disabled="!filtered || !usable" data-xs-slicer-clear @click="slicers.clear(ref_())">
        <IconFilterOff />
      </button>
    </div>
    <div class="xs-slicer-items" :style="style" role="listbox" :aria-multiselectable="true" :aria-label="slicer.caption ?? slicer.name">
      <button
        v-for="x in items"
        :key="x.item"
        type="button"
        role="option"
        :class="['xs-slicer-item', { 'xs-on': x.selected !== false, 'xs-nodata': mark && x.noData }]"
        :aria-selected="x.selected !== false"
        :disabled="!usable"
        :title="x.label"
        :data-xs-slicer-item="x.item"
        @click="pick($event, x.item)"
      >
        {{ x.label }}
      </button>
    </div>
  </div>
</template>
