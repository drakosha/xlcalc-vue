<script setup lang="ts">
// Excel's level buttons 1…N+1 of an outline bar, in the corner: the rows' buttons across the bar at the column letters'
// height, the columns' buttons down the bar above the row numbers. Level k collapses the groups at level k and deeper and expands the others.
import { computed } from 'vue'
import type { OutlineAxis } from '../core/outline'
import { useSheetContext } from './context'

const props = defineProps<{ axis: OutlineAxis }>()
const ctx = useSheetContext()
const ol = ctx.outline
const m = ctx.messages
const count = computed(() => (props.axis === 'rows' ? ctx.v.outlineSlots.value.rows : ctx.v.outlineSlots.value.cols))
</script>

<template>
  <th :class="['xs-ol-levels', `xs-ol-levels-${axis}`]" @click="ol.tap">
    <button
      v-for="k in count"
      :key="k"
      type="button"
      class="xs-ol-btn"
      :style="{ '--xs-ol-at': ol.at(k) + 'px' }"
      :disabled="!ol.can.value"
      :title="m.outlineLevel(k)"
      :aria-label="m.outlineLevel(k)"
      :data-xs-outline-level="k"
      @mousedown.prevent
      @click.stop="ol.showLevel(axis, k)"
    >
      {{ k }}
    </button>
  </th>
</template>
