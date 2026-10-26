<script setup lang="ts">
// Excel's Formula AutoComplete list below the editor (above it where the viewport has no room below), eight items at a time,
// each with the icon of its kind (a function, a defined name); the keys are the editor's (`useAutocomplete.onKey`), and a press keeps the editor's focus.
import { computed, ref } from 'vue'
import { FUNCTION_ROWS, type CompletionItem } from '../core/autocomplete'
import { useMessages } from '../messages'
import type { SheetMessages } from '../contract/messages'
import { IconFx } from './icons'
import { IconName } from './name-icons'
import { clampLeft, keepActiveInView, listHeight, placeBelowOrAbove, type PopupRect } from './popup'
import '../popup.css'

const props = withDefaults(
  defineProps<{
    /** The editor, in the coordinates of the element the list is placed in. */
    rect: PopupRect
    items: readonly CompletionItem[]
    /** The highlighted item. */
    active?: number
    /** The part in view, in the same coordinates: the list goes above the editor where it has no room below and stays inside horizontally. */
    view?: PopupRect
    /** Touch input: 44 px items, and a tap inserts one. */
    touch?: boolean
    messages?: Partial<SheetMessages>
  }>(),
  { active: 0, view: undefined, touch: false, messages: undefined },
)
const emit = defineEmits<{
  /** A mouse click: highlights the item. */
  highlight: [index: number]
  /** A double click, or a tap on a touch screen: inserts the item. */
  pick: [index: number]
}>()

const m = useMessages(() => props.messages)
/** The list is wide enough for the longest names (`ПРОМЕЖУТОЧНЫЕ.ИТОГИ`). */
const WIDTH_PX = 220
const listEl = ref<HTMLElement | null>(null)

const box = computed(() => {
  const r = props.rect
  const v = props.view
  const height = listHeight(props.items.length, FUNCTION_ROWS, props.touch)
  const width = v ? Math.min(WIDTH_PX, v.width) : WIDTH_PX
  return { top: `${placeBelowOrAbove(r, height, v)}px`, left: `${clampLeft(r.left, width, v)}px`, width: `${width}px`, maxHeight: `${height}px` }
})
keepActiveInView(listEl, () => [props.active, props.items], () => props.active)
</script>

<template>
  <ul
    ref="listEl"
    :class="['xs-popup-list', 'xs-fl', { 'xs-popup-touch xs-fl-touch': touch }]"
    role="listbox"
    :aria-label="m.functionList"
    :title="touch ? undefined : m.functionHint"
    :style="box"
    data-xs-functions
    @mousedown.prevent
  >
    <li
      v-for="(item, i) in items"
      :key="i"
      :class="['xs-popup-item', 'xs-fl-item', { 'xs-popup-on xs-fl-on': i === active }]"
      role="option"
      :aria-selected="i === active"
      :aria-label="item.kind === 'name' ? m.nameItem(item.text) : undefined"
      :data-xs-function="item.kind === 'function' ? item.text : undefined"
      :data-xs-name="item.kind === 'name' ? item.text : undefined"
      @click="touch ? emit('pick', i) : emit('highlight', i)"
      @dblclick="emit('pick', i)"
    >
      <IconName v-if="item.kind === 'name'" class="xs-fl-icon" />
      <IconFx v-else class="xs-fl-icon" />
      {{ item.text }}
    </li>
  </ul>
</template>

<style src="../autocomplete.css"></style>
