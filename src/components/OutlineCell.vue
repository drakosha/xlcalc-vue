<script setup lang="ts">
// A row's or column's cell in an outline bar (rows to the left of the numbers, columns above the letters): the brackets of the
// groups it is in and the buttons of the groups whose summary it is, each level in its slot.
import { computed } from 'vue'
import type { OutlineAxis } from '../core/outline'
import { useSheetContext } from './context'

const props = defineProps<{ axis: OutlineAxis; line: number }>()
const ctx = useSheetContext()
const ol = ctx.outline
const m = ctx.messages
const marks = computed(() => ol.marks(props.axis, props.line))
/** The summaries are before their groups (above, to the left): the brackets turn the other way. */
const before = computed(() => ol.outline.value && !(props.axis === 'rows' ? ol.outline.value.summaryBelow : ol.outline.value.summaryRight))
</script>

<template>
  <th :class="['xs-ol', `xs-ol-${axis}`, before ? 'xs-ol-before' : '']" @click="ol.tap">
    <span v-for="mk in marks" :key="mk.level" class="xs-ol-mark" :style="{ '--xs-ol-at': ol.at(mk.level) + 'px' }">
      <span v-if="mk.bracket" class="xs-ol-line"></span>
      <span v-if="mk.bracket === 'start'" class="xs-ol-tick"></span>
      <template v-if="mk.button">
        <span v-if="!mk.button.collapsed" class="xs-ol-end"></span>
        <button
          type="button"
          :class="['xs-ol-btn', mk.button.collapsed ? 'xs-ol-plus' : 'xs-ol-minus']"
          :disabled="!ol.can.value"
          :title="mk.button.collapsed ? m.outlineExpand : m.outlineCollapse"
          :aria-label="mk.button.collapsed ? m.outlineExpand : m.outlineCollapse"
          :aria-expanded="!mk.button.collapsed"
          :data-xs-outline="`${mk.button.from}:${mk.button.to}:${mk.button.level}`"
          @mousedown.prevent
          @click.stop="ol.toggle(axis, mk.button)"
        ></button>
      </template>
    </span>
  </th>
</template>

<style src="../outline.css"></style>
