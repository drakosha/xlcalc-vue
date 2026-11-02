<script setup lang="ts">
// Excel's Font, Font Size and Number Format boxes on the toolbar: an input that shows the active cell's value and takes a
// typed one or a pick from its list (Number Format's only a pick). Enter or a pick applies; Esc, and leaving the box without
// Enter, restore what it showed.
import { computed, onBeforeUnmount, ref, useId, watch } from 'vue'
import { fontCss } from '../core/cell-style'
import { IconDown } from './icons'
import { keepActiveInView } from './popup'

const props = defineProps<{
  /** What the box shows while it is not being edited: the active cell's font or size. */
  value: string
  items: readonly string[]
  label: string
  /** The input's hook (`data-tb`); the button that opens the list has `<hook>-list`. */
  hook: string
  /** The list is open: the toolbar keeps one menu open at a time. */
  open: boolean
  disabled?: boolean
  /** The input's width class. */
  fieldClass?: string
  /** Each item in its own font, as Excel's Font list draws them. */
  preview?: boolean
  /** Typing completes the text with the first item it begins, as Excel's Font box does. */
  complete?: boolean
  /** Why a text cannot apply, or null; the box then keeps the text, marked invalid, until it is corrected or Esc restores it. */
  check?: (text: string) => string | null
  inputmode?: 'text' | 'decimal'
  /** The field only shows its value, as the Number Format box takes no typing: a click opens the list, the arrows go through it, Enter picks. */
  readonly?: boolean
  /** Each item's key, its hook then `data-tb-<hook>="<key>"` in place of `data-tb-item`. */
  keys?: readonly string[]
  /** The texts the field shows, which it is as wide as the widest of, in place of a width class. */
  fit?: readonly string[]
  /** A command after a line at the list's end (More Number Formats...), its hook `<key>` as the items'; it emits `more`. */
  more?: { label: string; key: string }
  /** The list shows every item without scrolling, as Excel's Number Format list (on a touch screen it still ends within the viewport). */
  whole?: boolean
  /** The item marked as the list opens, -1 for none, in place of the one the box's value names (Number Format's by the cell's code). */
  marked?: number
}>()
// the box's other attributes (hooks) go on its field
defineOptions({ inheritAttrs: false })
const emit = defineEmits<{
  /** A text applies: an item's in its own spelling (with its index), or a typed one (index -1). */
  apply: [text: string, index: number]
  'update:open': [open: boolean]
  /** Esc left the box, as Excel returns to the sheet. */
  leave: []
  more: []
}>()

const id = useId()
const input = ref<HTMLInputElement | null>(null)
const listEl = ref<HTMLElement | null>(null)
/** The text being typed, null while the box shows the active cell's value. */
const draft = ref<string | null>(null)
const bad = ref<string | null>(null)
const active = ref(-1)
const shown = computed(() => draft.value ?? props.value)
const indexOf = (text: string) => props.items.findIndex((x) => x.toLowerCase() === text.trim().toLowerCase())
const itemFont = (name: string) => (props.preview ? { fontFamily: fontCss({ font: { name } }, undefined).fontFamily } : undefined)
/** The list's child of an item: the command comes after the line before it. */
const childOf = (i: number) => (props.more && i === props.items.length ? i + 1 : i)
keepActiveInView(listEl, () => [props.open, active.value], () => childOf(active.value))
/** An option's id, for the field's `aria-activedescendant`: a screen reader names the item the arrows reach. */
const optionId = (i: number) => `${id}-o${i}`
/**
 * The room from the list's top to the viewport's bottom, which its height never passes (a phone's More menu opens it low).
 * Measured as it opens and again on the window's resize and on any scroll while it is open, so a rotated phone or a scrolled
 * page keeps it within the window: recomputed rather than closed, as a scroll inside the list itself must not close it.
 */
const room = ref<string | undefined>(undefined)
function measureRoom(): void {
  const top = props.open ? listEl.value?.getBoundingClientRect().top : undefined
  room.value = top === undefined ? undefined : `${Math.max(120, window.innerHeight - top - 8)}px`
}
watch(
  () => props.open,
  (on) => {
    measureRoom()
    if (typeof window === 'undefined') return
    const how = { capture: true, passive: true }
    if (on) {
      window.addEventListener('resize', measureRoom)
      window.addEventListener('scroll', measureRoom, how)
    } else stopRoom()
  },
  { flush: 'post' },
)
function stopRoom(): void {
  window.removeEventListener('resize', measureRoom)
  window.removeEventListener('scroll', measureRoom, { capture: true })
}
onBeforeUnmount(() => typeof window !== 'undefined' && stopRoom())

function setOpen(on: boolean): void {
  if (on) active.value = props.marked ?? indexOf(shown.value)
  if (on !== props.open) emit('update:open', on)
}
function restore(): void {
  draft.value = null
  bad.value = null
  setOpen(false)
}

// A click selects the whole text, as in Excel's boxes: the mouseup after the focus would otherwise drop the selection.
let focused = false
function onFocus(): void {
  focused = true
  input.value?.select()
}
function onMouseUp(e: MouseEvent): void {
  if (focused) e.preventDefault()
  focused = false
}

function onInput(e: Event): void {
  const el = e.target as HTMLInputElement
  const typed = el.value
  bad.value = null
  const at = props.complete && typed && (e as InputEvent).inputType?.startsWith('insert') ? props.items.findIndex((x) => x.toLowerCase().startsWith(typed.toLowerCase())) : -1
  if (at >= 0) {
    el.value = props.items[at]!
    el.setSelectionRange(typed.length, el.value.length)
  }
  draft.value = el.value
  active.value = at >= 0 ? at : indexOf(el.value)
}

/** Applies a text: an item of the list in its own spelling; an empty text restores the box, an invalid one stays. */
function commit(text: string): void {
  const t = text.trim()
  if (!t) return restore()
  const why = props.check?.(t) ?? null
  if (why) {
    bad.value = why
    input.value?.select()
    return
  }
  const at = indexOf(t)
  restore()
  emit('apply', at >= 0 ? props.items[at]! : t, at)
  input.value?.blur()
}

/** The list's command after its items: it closes the list and leaves the box as Esc does. */
function runMore(): void {
  restore()
  input.value?.blur()
  emit('more')
}

/** The arrows go through the list (its command too), putting each item into a box that takes typing; Alt+↓ or an arrow on a closed list opens it. */
function move(by: number, open: boolean): void {
  if (!props.open || open) return setOpen(true)
  const at = Math.max(0, Math.min(props.items.length - (props.more ? 0 : 1), active.value + by))
  active.value = at
  if (!props.readonly) draft.value = props.items[at] ?? draft.value
  bad.value = null
}
/** Enter in a box without typing: opens the list, then picks the item the arrows reached; with none marked it leaves for the grid, as Esc. */
function pick(): void {
  if (!props.open) return setOpen(true)
  if (props.more && active.value === props.items.length) return runMore()
  const item = props.items[active.value]
  if (item !== undefined) return commit(item)
  restore()
  input.value?.blur()
  emit('leave')
}
const itemHook = (x: string, i: number, key = props.keys?.[i]) => (key === undefined ? { 'data-tb-item': x } : { [`data-tb-${props.hook}`]: key })

function onKey(e: KeyboardEvent): void {
  if (e.isComposing) return
  if (e.key === 'Enter') {
    e.preventDefault()
    if (props.readonly) pick()
    else commit(input.value?.value ?? '')
  } else if (e.key === 'Escape') {
    e.preventDefault()
    restore()
    input.value?.blur()
    emit('leave')
  } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault()
    move(e.key === 'ArrowDown' ? 1 : -1, e.altKey)
    if (draft.value !== null) input.value?.select()
  }
}
</script>

<template>
  <span class="xs-tb-split xs-tb-box">
    <span :class="['xs-tb-field', { 'xs-tb-fit': fit }]">
      <input
        ref="input"
        v-bind="$attrs"
        :class="['xs-field', 'xs-field-s', 'xs-tb-input', fieldClass]"
        type="text"
        role="combobox"
        spellcheck="false"
        autocomplete="off"
        :readonly="readonly"
        :inputmode="inputmode"
        :data-tb="hook"
        :value="shown"
        :data-tip="bad ?? label"
        :aria-label="label"
        :aria-expanded="open"
        :aria-controls="id"
        :aria-activedescendant="open && active >= 0 ? optionId(active) : undefined"
        :aria-invalid="!!bad"
        :disabled="disabled"
        @focus="onFocus"
        @mouseup="onMouseUp"
        @click="readonly && setOpen(!open)"
        @input="onInput"
        @keydown="onKey"
        @blur="restore"
      />
      <span v-for="x in fit" :key="x" class="xs-field xs-field-s xs-tb-input xs-tb-fit-text" aria-hidden="true">{{ x }}</span>
    </span>
    <button type="button" class="xs-tool xs-tb-caret" :data-tb="`${hook}-list`" :data-tip="label" :aria-label="label" :aria-expanded="open" :disabled="disabled" @click="setOpen(!open)">
      <IconDown />
    </button>
    <div v-if="open" :id="id" ref="listEl" :class="['xs-tb-pop', 'xs-tb-list', { 'xs-tb-whole': whole }]" role="listbox" :aria-label="label" :style="{ '--xs-tb-room': room }">
      <button
        v-for="(x, i) in items"
        :key="x"
        type="button"
        class="xs-tb-item"
        role="option"
        :aria-selected="i === active"
        :id="optionId(i)"
        v-bind="itemHook(x, i)"
        :style="itemFont(x)"
        @click="commit(x)"
      >
        {{ x }}
      </button>
      <template v-if="more">
        <div class="xs-tb-sep" role="presentation" />
        <button type="button" class="xs-tb-item" role="option" :aria-selected="active === items.length" :id="optionId(items.length)" v-bind="itemHook(more.label, -1, more.key)" @click="runMore">
          {{ more.label }}
        </button>
      </template>
    </div>
  </span>
</template>
