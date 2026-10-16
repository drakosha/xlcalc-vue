<script setup lang="ts">
// Excel's Formula AutoComplete list under the field being edited (over it where the view has no room below), eight
// functions at a time; the keys are the field's (`useAutocomplete.onKey`) and a press keeps the field's focus.
import { computed, ref } from 'vue'
import { FUNCTION_ROWS } from '../core/autocomplete'
import { useMessages, type SheetMessages } from '../messages'
import { clampLeft, keepActiveInView, listHeight, placeBelowOrAbove, type PopupRect } from './popup'
import '../popup.css'

const props = withDefaults(
  defineProps<{
    /** The field, in the coordinates of the element the list is placed in. */
    rect: PopupRect
    items: readonly string[]
    /** The highlighted item. */
    active?: number
    /** What is in view, in the same coordinates: the list goes over the field where it has no room below and stays inside it sideways. */
    view?: PopupRect
    /** A finger: items of 44 px, a tap enters one. */
    touch?: boolean
    messages?: Partial<SheetMessages>
  }>(),
  { active: 0, view: undefined, touch: false, messages: undefined },
)
const emit = defineEmits<{
  /** A click with the mouse: highlight the item. */
  highlight: [index: number]
  /** A double click, or a tap under a finger: enter the item. */
  pick: [index: number]
}>()

const m = useMessages(() => props.messages)
/** The list's width holds the longest names (`ПРОМЕЖУТОЧНЫЕ.ИТОГИ`). */
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
      v-for="(name, i) in items"
      :key="i"
      :class="['xs-popup-item', 'xs-fl-item', { 'xs-popup-on xs-fl-on': i === active }]"
      role="option"
      :aria-selected="i === active"
      :data-xs-function="name"
      @click="touch ? emit('pick', i) : emit('highlight', i)"
      @dblclick="emit('pick', i)"
    >
      {{ name }}
    </li>
  </ul>
</template>

<style src="../autocomplete.css"></style>
