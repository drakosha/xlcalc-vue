<script setup lang="ts">
// The context menus of rows, columns, tabs and cells at the pointer (kept on screen) and their small dialogs: a size in Excel's units,
// the confirmation before a sheet with data is deleted, the hidden sheets to unhide, rows or columns to group. A click outside or Esc closes them.
// Keys in a menu as in Excel: Down and Up cycle through its enabled items, Home and End, Enter or Space chooses, Esc and Tab close;
// a press in the menu leaves focus where the keys put it.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, useId, watch } from 'vue'
import type { MenuItem } from '../core/menu'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { useOutsidePress } from './popup'
import '../controls.css'

const ctx = useSheetContext()
const { menu } = ctx
const m = ctx.messages

/** The margin kept between the menu and the window's edge. */
const EDGE_PX = 4

const box = ref<HTMLElement | null>(null)
const pos = ref({ left: '0px', top: '0px' })
/** Where focus was before the menu opened, to return to when the menu closes with focus in it. */
let before: HTMLElement | null = null
// Opens at the pointer, then is moved inside the window: to the left of a right edge, above a pointer near the bottom (the tabs).
watch(menu.open, async (o, was) => {
  if (!o) return void (was && giveBack())
  if (!was) before = document.activeElement instanceof HTMLElement ? document.activeElement : null
  pos.value = { left: `${o.x}px`, top: `${o.y}px` }
  await nextTick()
  const r = box.value?.getBoundingClientRect()
  if (!r) return
  const left = Math.max(EDGE_PX, Math.min(o.x, window.innerWidth - r.width - EDGE_PX))
  const top = o.y + r.height > window.innerHeight - EDGE_PX ? Math.max(EDGE_PX, o.y - r.height) : o.y
  pos.value = { left: `${left}px`, top: `${top}px` }
  enabled()[0]?.focus({ preventScroll: true })
})
/** Returns focus to the element focused before the menu (the grid if it is gone), when focus is still in the menu. */
function giveBack(): void {
  const el = before
  before = null
  if (!box.value?.contains(document.activeElement)) return
  if (el?.isConnected) el.focus({ preventScroll: true })
  else ctx.focusGrid()
}

const enabled = () => [...(box.value?.querySelectorAll<HTMLButtonElement>('[role=menuitem]:not(:disabled)') ?? [])]
function onMenuKey(e: KeyboardEvent): void {
  const list = enabled()
  const at = list.indexOf(document.activeElement as HTMLButtonElement)
  let next: number
  if (e.key === 'ArrowDown') next = (at + 1) % list.length
  else if (e.key === 'ArrowUp') next = at <= 0 ? list.length - 1 : at - 1
  else if (e.key === 'Home') next = 0
  else if (e.key === 'End') next = list.length - 1
  else if (e.key === 'Enter' || e.key === ' ' || e.key === 'Tab') next = -1
  else return
  // the menu's own keys: no page scroll, no button click after ours, nothing for the grid or the page
  e.preventDefault()
  e.stopPropagation()
  if (next >= 0) return void list[next]?.focus({ preventScroll: true })
  const item = e.key === 'Tab' ? undefined : menu.items.value.find((x) => x.id === list[at]?.dataset.xsMenu)
  if (item) menu.choose(item)
  else menu.close()
}

/** Whether the pointer went down in the menu: a click without that is the release of the long press that opened the menu over the finger, not a choice. */
let pressedIn = false
function onItemClick(e: MouseEvent, item: MenuItem): void {
  const chosen = e.detail === 0 || pressedIn
  pressedIn = false
  if (chosen) menu.choose(item)
}

const size = computed(() => (menu.dialog.value?.kind === 'size' ? menu.dialog.value : null))
const doomed = computed(() => (menu.dialog.value?.kind === 'deleteSheet' ? menu.dialog.value.sheet : null))
const hidden = computed(() => (menu.dialog.value?.kind === 'unhide' ? menu.dialog.value.sheets : null))
const grouping = computed(() => (menu.dialog.value?.kind === 'outline' ? menu.dialog.value : null))
/** A size dialog opens with its input's text selected, a grouping dialog with its first choice. */
const dialogFocus = computed(() => (size.value ? '[data-xs-dialog-input]' : grouping.value ? '[data-xs-outline-axis]' : undefined))
const closeDialog = () => (menu.dialog.value = null)
/** The dialog's title labels it for screen readers. */
const titleId = `${useId()}-title`

useOutsidePress(box, () => menu.close(), () => !!menu.open.value)
function onDocKey(e: KeyboardEvent): void {
  if (e.key !== 'Escape' || (!menu.open.value && !menu.dialog.value)) return
  e.preventDefault()
  menu.close()
  closeDialog()
}
onMounted(() => document.addEventListener('keydown', onDocKey))
onBeforeUnmount(() => document.removeEventListener('keydown', onDocKey))
</script>

<template>
  <div v-if="menu.open.value" ref="box" class="xs-menu" role="menu" :aria-label="menu.open.value.kind === 'cells' ? m.cellMenu : m.menu" :style="pos" data-xs-menu-box @keydown="onMenuKey" @mousedown.prevent @pointerdown="pressedIn = true">
    <div class="xs-menu-caption">{{ menu.caption.value }}</div>
    <template v-for="item in menu.items.value" :key="item.id">
      <div v-if="item.group" class="xs-menu-sep" role="separator"></div>
      <button type="button" role="menuitem" tabindex="-1" class="xs-menu-item" :data-xs-menu="item.id" :disabled="item.disabled" @click="onItemClick($event, item)">{{ item.label }}</button>
    </template>
  </div>
  <DialogShell v-if="menu.dialog.value" class="xs-dialog" :focus="dialogFocus" :aria-labelledby="titleId" data-xs-dialog @close="closeDialog">
    <template v-if="size">
      <label class="xs-dialog-field">
        <span :id="titleId" class="xs-modal-title">{{ size.axis === 'columns' ? m.columnWidth : m.rowHeight }}</span>
        <input v-model="size.draft" class="xs-field xs-field-l" type="text" inputmode="decimal" autocomplete="off" spellcheck="false" data-xs-dialog-input @keydown.enter.prevent="menu.applySize()" />
      </label>
      <p v-if="size.error" class="xs-dialog-error" role="alert">{{ size.error }}</p>
    </template>
    <template v-else-if="doomed !== null">
      <p :id="titleId" class="xs-modal-title">{{ m.deleteSheet }}</p>
      <p>{{ m.deleteSheetText(doomed) }}</p>
    </template>
    <template v-else-if="grouping">
      <p :id="titleId" class="xs-modal-title">{{ grouping.op === 'group' ? m.menuGroup : m.menuUngroup }}</p>
      <div class="xs-dialog-choices" role="radiogroup" @keydown.enter.prevent="menu.applyOutline()">
        <label class="xs-dialog-choice"><input v-model="grouping.axis" type="radio" value="rows" data-xs-outline-axis="rows" />{{ m.outlineRows }}</label>
        <label class="xs-dialog-choice"><input v-model="grouping.axis" type="radio" value="columns" data-xs-outline-axis="columns" />{{ m.outlineColumns }}</label>
      </div>
    </template>
    <template v-else-if="hidden">
      <p :id="titleId" class="xs-modal-title">{{ m.unhideSheet }}</p>
      <div class="xs-dialog-list">
        <button v-for="name in hidden" :key="name" type="button" class="xs-menu-item" :data-xs-unhide="name" @click="menu.unhide(name)">{{ name }}</button>
      </div>
    </template>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="closeDialog">{{ m.cancel }}</button>
      <button v-if="size" type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="menu.applySize()">{{ m.ok }}</button>
      <button v-else-if="grouping" type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="menu.applyOutline()">{{ m.ok }}</button>
      <button v-else-if="doomed !== null" type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="menu.confirmDelete()">{{ m.menuDelete }}</button>
    </div>
  </DialogShell>
</template>

<style src="../menu.css"></style>
