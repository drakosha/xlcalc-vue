<script setup lang="ts">
// Excel's dropdown of a list rule: the arrow right of the active cell and the list of items under it (over it where the
// view has no room below), eight at a time; the keys are the sheet's (`useValidation.onKey`): a press keeps the grid's focus.
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import { LIST_ROWS, type ValidationRect } from '../core/validation'
import { useMessages, type SheetMessages } from '../messages'

const props = withDefaults(
  defineProps<{
    /** The active cell (a merged area whole). */
    rect: ValidationRect
    items: string[]
    open: boolean
    /** The highlighted item; -1 for none. */
    active?: number
    /** What of the grid is in view, in the same coordinates: the list goes over the cell where it has no room below. */
    view?: ValidationRect
    /** A finger: the arrow and the items are 44 px. */
    touch?: boolean
    messages?: Partial<SheetMessages>
  }>(),
  { active: -1, view: undefined, touch: false, messages: undefined },
)
const emit = defineEmits<{
  /** The arrow: open or close the list. */
  toggle: []
  /** An item clicked: enter its text into the cell. */
  pick: [index: number, text: string]
  /** The pointer over an item: highlight it. */
  hover: [index: number]
  /** A press outside the dropdown while the list is open. */
  close: []
}>()

const m = useMessages(() => props.messages)
/** Excel's sizes: the arrow 17 × 18 px beside the cell, items of 18 px; 44 under a finger. */
const ARROW_PX = 17
const ITEM_PX = 18
const TOUCH_PX = 44
const root = ref<HTMLElement | null>(null)
const listEl = ref<HTMLElement | null>(null)

const arrowW = computed(() => (props.touch ? TOUCH_PX : ARROW_PX))
const anchor = computed(() => ({ top: `${props.rect.top}px`, left: `${props.rect.left + props.rect.width}px`, height: `${props.rect.height}px` }))
/** The list as wide as the cell and its arrow, under the cell, or over it when the view has room above alone. */
const box = computed(() => {
  const r = props.rect
  const height = Math.min(props.items.length, LIST_ROWS) * (props.touch ? TOUCH_PX : ITEM_PX) + 2
  const below = r.top + r.height
  const v = props.view
  const up = !!v && below + height > v.top + v.height && r.top - height >= v.top
  return { left: `${r.left}px`, top: `${up ? r.top - height : below}px`, width: `${r.width + arrowW.value}px`, maxHeight: `${height}px` }
})

// The highlighted item stays in view as the keys move it.
watch(
  () => [props.open, props.active] as const,
  async ([open, i]) => {
    if (!open || i < 0) return
    await nextTick()
    listEl.value?.children[i]?.scrollIntoView?.({ block: 'nearest' })
  },
)
function onOutside(e: PointerEvent): void {
  if (props.open && !root.value?.contains(e.target as Node)) emit('close')
}
watch(
  () => props.open,
  (open) => (open ? document.addEventListener('pointerdown', onOutside, true) : document.removeEventListener('pointerdown', onOutside, true)),
  { immediate: true },
)
onBeforeUnmount(() => document.removeEventListener('pointerdown', onOutside, true))
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
    <ul v-if="open" ref="listEl" class="xs-vd-list" role="listbox" :aria-label="m.validationList" :style="box" data-xs-dropdown-list>
      <li
        v-for="(item, i) in items"
        :key="i"
        :class="['xs-vd-item', { 'xs-vd-on': i === active }]"
        role="option"
        :aria-selected="i === active"
        :title="item"
        :data-xs-dropdown-item="i"
        @pointermove="i !== active && emit('hover', i)"
        @click="emit('pick', i, item)"
      >
        {{ item || ' ' }}
      </li>
    </ul>
  </div>
</template>

<style src="../validation.css"></style>
