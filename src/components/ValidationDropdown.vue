<script setup lang="ts">
// Excel's dropdown for a list rule: the arrow to the right of the active cell and the item list below it (above it where the
// viewport has no room below), eight at a time; the keys are the sheet's (`useValidation.onKey`): a press keeps the grid's focus.
import { computed, ref } from 'vue'
import { LIST_ROWS } from '../core/validation'
import { useMessages } from '../messages'
import type { SheetMessages } from '../contract/messages'
import { keepActiveInView, listHeight, placeBelowOrAbove, TOUCH_PX, useOutsidePress, type PopupRect } from './popup'
import '../popup.css'

const props = withDefaults(
  defineProps<{
    /** The active cell (an entire merged range). */
    rect: PopupRect
    items: string[]
    open: boolean
    /** The highlighted item; -1 for none. */
    active?: number
    /** The part of the grid in view, in the same coordinates: the list goes above the cell where it has no room below. */
    view?: PopupRect
    /** Touch input: the arrow and the items are 44 px. */
    touch?: boolean
    messages?: Partial<SheetMessages>
  }>(),
  { active: -1, view: undefined, touch: false, messages: undefined },
)
const emit = defineEmits<{
  /** The arrow: opens or closes the list. */
  toggle: []
  /** An item was clicked: enters it into the cell. */
  pick: [index: number]
  /** The pointer is over an item: highlights it. */
  hover: [index: number]
  /** A press outside the dropdown while the list is open. */
  close: []
}>()

const m = useMessages(() => props.messages)
/** Excel's arrow is 17 × 18 px next to the cell; 44 on a touch screen. */
const ARROW_PX = 17
const root = ref<HTMLElement | null>(null)
const listEl = ref<HTMLElement | null>(null)

const arrowW = computed(() => (props.touch ? TOUCH_PX : ARROW_PX))
const anchor = computed(() => ({ top: `${props.rect.top}px`, left: `${props.rect.left + props.rect.width}px`, height: `${props.rect.height}px` }))
/** The list, as wide as the cell plus its arrow, below the cell, or above it when the viewport has room only above. */
const box = computed(() => {
  const r = props.rect
  const height = listHeight(props.items.length, LIST_ROWS, props.touch)
  return { left: `${r.left}px`, top: `${placeBelowOrAbove(r, height, props.view)}px`, width: `${r.width + arrowW.value}px`, maxHeight: `${height}px` }
})
keepActiveInView(listEl, () => [props.open, props.active], () => props.active)
useOutsidePress(root, () => emit('close'), () => props.open)
</script>

<template>
  <div ref="root" :class="['xs-vd', { 'xs-vd-touch': touch }]" data-xs-dropdown @mousedown.prevent>
    <div class="xs-vd-anchor" :style="anchor">
      <button
        type="button"
        class="xs-vd-arrow"
        tabindex="-1"
        :title="m.validationList"
        :aria-label="m.validationList"
        aria-haspopup="listbox"
        :aria-expanded="open"
        data-xs-dropdown-arrow
        @click="emit('toggle')"
      >
        <svg viewBox="0 0 10 10" aria-hidden="true"><path d="M2 3.5h6L5 7z" fill="currentColor" /></svg>
      </button>
    </div>
    <ul v-if="open" ref="listEl" :class="['xs-popup-list', 'xs-vd-list', { 'xs-popup-touch': touch }]" role="listbox" :aria-label="m.validationList" :style="box" data-xs-dropdown-list>
      <li
        v-for="(item, i) in items"
        :key="i"
        :class="['xs-popup-item', 'xs-vd-item', { 'xs-popup-on': i === active }]"
        role="option"
        :aria-selected="i === active"
        :title="item"
        :data-xs-dropdown-item="i"
        @pointermove="i !== active && emit('hover', i)"
        @click="emit('pick', i)"
      >
        {{ item || ' ' }}
      </li>
    </ul>
  </div>
</template>

<style src="../validation.css"></style>
