<script setup lang="ts">
// Excel's Name Box at the left of the formula bar: the selection's address, or the name whose range it is; a typed address or
// name goes there (a new name defines one over the selection), and its arrow lists the names that refer to ranges, in alphabetical order.
// Keys in the box: Enter goes, Esc gives the box back, Alt+Down opens the list, where Down and Up move and Enter goes.
import { ref } from 'vue'
import type { SheetName } from '../contract/workbook'
import { listStep, stepIndex } from '../core/list-keys'
import { useSheetContext } from './context'
import { IconDown } from './icons'
import { keepActiveInView, listHeight, useOutsidePress } from './popup'
import '../popup.css'
import '../names.css'

const ctx = useSheetContext()
const { nav, names } = ctx
const m = ctx.messages

/** The names the list shows at once, as Excel's eight. */
const ROWS = 8
const box = ref<HTMLElement | null>(null)
const listEl = ref<HTMLElement | null>(null)
const open = ref(false)
/** The highlighted name; -1 for none. */
const index = ref(-1)

function toggle(): void {
  if (open.value) return close()
  open.value = true
  index.value = names.boxNames.value.findIndex((n) => n.name === nav.boxText.value)
}
function close(): void {
  open.value = false
  index.value = -1
}
/** A name picked from the list: its range is selected, and the grid takes focus. */
async function pick(n: SheetName): Promise<void> {
  close()
  nav.addressFocused.value = false
  await nav.goName(n)
}
useOutsidePress(box, close, () => open.value)
keepActiveInView(listEl, index, () => index.value)

function onFocus(e: FocusEvent): void {
  nav.addressFocused.value = true
  ;(e.target as HTMLInputElement).select()
}
function onBlur(): void {
  nav.addressFocused.value = false
  nav.addressDraft.value = nav.boxText.value
}
async function onEnter(e: KeyboardEvent): Promise<void> {
  const input = e.target as HTMLInputElement
  const chosen = open.value ? names.boxNames.value[index.value] : undefined
  if (chosen) {
    input.blur()
    return pick(chosen)
  }
  close()
  await nav.goAddress()
  input.blur()
}
function onEsc(e: KeyboardEvent): void {
  if (open.value) return close()
  nav.addressDraft.value = nav.boxText.value
  ;(e.target as HTMLInputElement).blur()
}
function onKey(e: KeyboardEvent): void {
  if (e.altKey && e.key === 'ArrowDown' && names.available.value) {
    e.preventDefault()
    return toggle()
  }
  const count = names.boxNames.value.length
  const by = open.value && count && !e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey ? listStep(e.key, count, ROWS) : null
  if (by === null) return
  e.preventDefault()
  index.value = stepIndex(index.value, by, count)
}
</script>

<template>
  <div ref="box" class="xs-namebox">
    <input
      v-model="nav.addressDraft.value"
      class="xs-address"
      type="text"
      spellcheck="false"
      placeholder="A1"
      :title="m.address"
      :aria-expanded="names.available.value ? (open ? 'true' : 'false') : undefined"
      @focus="onFocus"
      @blur="onBlur"
      @keydown.enter.prevent="onEnter"
      @keydown.esc.prevent="onEsc"
      @keydown="onKey"
    />
    <button
      v-if="names.available.value"
      type="button"
      class="xs-tool xs-namebox-arrow"
      :title="m.nameBoxList"
      :aria-label="m.nameBoxList"
      :aria-expanded="open ? 'true' : 'false'"
      data-xs-namebox-arrow
      @mousedown.prevent
      @click="toggle"
    >
      <IconDown />
    </button>
    <ul
      v-if="open && names.boxNames.value.length"
      ref="listEl"
      :class="['xs-popup-list', 'xs-namebox-list', { 'xs-popup-touch': ctx.v.touch.value }]"
      role="listbox"
      :aria-label="m.nameBoxList"
      :style="{ maxHeight: `${listHeight(names.boxNames.value.length, ROWS, ctx.v.touch.value)}px` }"
      data-xs-namebox-list
      @mousedown.prevent
    >
      <li
        v-for="(n, i) in names.boxNames.value"
        :key="`${n.scope ?? ''}!${n.name}`"
        :class="['xs-popup-item', 'xs-namebox-item', { 'xs-popup-on': i === index }]"
        role="option"
        :aria-selected="i === index"
        :data-xs-namebox-name="n.name"
        @mouseenter="index = i"
        @click="pick(n)"
      >
        {{ n.name }}
      </li>
    </ul>
  </div>
</template>
