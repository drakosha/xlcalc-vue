<script setup lang="ts">
// Excel's Formula AutoComplete list under the field being edited (over it where the view has no room below), eight
// functions at a time; the keys are the field's (`useAutocomplete.onKey`) and a press keeps the field's focus.
import { computed, nextTick, ref, watch } from 'vue'
import { FUNCTION_ROWS } from '../core/autocomplete'
import type { ValidationRect } from '../core/validation'
import { enMessages, type SheetMessages } from '../messages'

const props = withDefaults(
  defineProps<{
    /** The field, in the coordinates of the element the list is placed in. */
    rect: ValidationRect
    items: readonly string[]
    /** The highlighted item. */
    active?: number
    /** What is in view, in the same coordinates: the list goes over the field where it has no room below and stays inside it sideways. */
    view?: ValidationRect
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

const m = computed<SheetMessages>(() => ({ ...enMessages, ...props.messages }))
/** Items of 18 px as Excel's, 44 under a finger; the list's width holds the longest names (`ПРОМЕЖУТОЧНЫЕ.ИТОГИ`). */
const ITEM_PX = 18
const TOUCH_PX = 44
const WIDTH_PX = 220
const listEl = ref<HTMLElement | null>(null)

const box = computed(() => {
  const r = props.rect
  const v = props.view
  const height = Math.min(props.items.length, FUNCTION_ROWS) * (props.touch ? TOUCH_PX : ITEM_PX) + 2
  const width = v ? Math.min(WIDTH_PX, v.width) : WIDTH_PX
  const below = r.top + r.height
  const up = !!v && below + height > v.top + v.height && r.top - height >= v.top
  const left = v ? Math.max(v.left, Math.min(r.left, v.left + v.width - width)) : r.left
  return { top: `${up ? r.top - height : below}px`, left: `${left}px`, width: `${width}px`, maxHeight: `${height}px` }
})

// The highlighted item stays in view as the keys move it, by the list's own scroll alone.
watch(
  () => [props.active, props.items] as const,
  async ([i]) => {
    await nextTick()
    const list = listEl.value
    const item = list?.children[i] as HTMLElement | undefined
    if (!list || !item) return
    if (item.offsetTop < list.scrollTop) list.scrollTop = item.offsetTop
    else if (item.offsetTop + item.offsetHeight > list.scrollTop + list.clientHeight) list.scrollTop = item.offsetTop + item.offsetHeight - list.clientHeight
  },
)
</script>

<template>
  <ul
    ref="listEl"
    :class="['xs-fl', { 'xs-fl-touch': touch }]"
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
      :class="['xs-fl-item', { 'xs-fl-on': i === active }]"
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
