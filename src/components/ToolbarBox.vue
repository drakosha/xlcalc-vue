<script setup lang="ts">
// Excel's Font and Font Size boxes on the toolbar: an input that shows the active cell's value and takes a typed one or a
// pick from its list. Enter or a pick applies; Esc, and leaving the box without Enter, restore what it showed.
import { computed, ref, useId } from 'vue'
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
}>()
const emit = defineEmits<{
  apply: [text: string]
  'update:open': [open: boolean]
  /** Esc left the box, as Excel returns to the sheet. */
  leave: []
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
keepActiveInView(listEl, () => [props.open, active.value], () => active.value)

function setOpen(on: boolean): void {
  if (on) active.value = indexOf(shown.value)
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
  emit('apply', at >= 0 ? props.items[at]! : t)
  input.value?.blur()
}

/** The arrows go through the list, putting each item into the box; Alt+↓ or an arrow on a closed list opens it. */
function move(by: number, open: boolean): void {
  if (!props.open || open) return setOpen(true)
  const at = Math.max(0, Math.min(props.items.length - 1, active.value + by))
  active.value = at
  draft.value = props.items[at] ?? draft.value
  bad.value = null
}

function onKey(e: KeyboardEvent): void {
  if (e.isComposing) return
  if (e.key === 'Enter') {
    e.preventDefault()
    commit(input.value?.value ?? '')
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
    <input
      ref="input"
      :class="['xs-field', 'xs-field-s', 'xs-tb-input', fieldClass]"
      type="text"
      role="combobox"
      spellcheck="false"
      autocomplete="off"
      :inputmode="inputmode"
      :data-tb="hook"
      :value="shown"
      :data-tip="bad ?? label"
      :aria-label="label"
      :aria-expanded="open"
      :aria-controls="id"
      :aria-invalid="!!bad"
      :disabled="disabled"
      @focus="onFocus"
      @mouseup="onMouseUp"
      @input="onInput"
      @keydown="onKey"
      @blur="restore"
    />
    <button type="button" class="xs-tool xs-tb-caret" :data-tb="`${hook}-list`" :data-tip="label" :aria-label="label" :aria-expanded="open" :disabled="disabled" @click="setOpen(!open)">
      <IconDown />
    </button>
    <div v-if="open" :id="id" ref="listEl" class="xs-tb-pop xs-tb-list" role="listbox" :aria-label="label">
      <button
        v-for="(x, i) in items"
        :key="x"
        type="button"
        class="xs-tb-item"
        role="option"
        :aria-selected="i === active"
        :data-tb-item="x"
        :style="itemFont(x)"
        @click="commit(x)"
      >
        {{ x }}
      </button>
    </div>
  </span>
</template>
