<script setup lang="ts">
// The menus of lines, tabs and cells at the pointer (kept on screen) and their small dialogs: a size in Excel's units, the question before a
// sheet with data is deleted, the hidden sheets to show, rows or columns to group. A press outside or Esc closes them.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { useSheetContext } from './context'

const ctx = useSheetContext()
const { menu } = ctx
const m = ctx.messages

/** Room kept between the menu and the window's edge. */
const EDGE_PX = 4

const box = ref<HTMLElement | null>(null)
const pos = ref({ left: '0px', top: '0px' })
// Open at the pointer, then moved inside the window: left of a right edge, above a pointer near the bottom (the tabs).
watch(menu.open, async (o) => {
  if (!o) return
  pos.value = { left: `${o.x}px`, top: `${o.y}px` }
  await nextTick()
  const r = box.value?.getBoundingClientRect()
  if (!r) return
  const left = Math.max(EDGE_PX, Math.min(o.x, window.innerWidth - r.width - EDGE_PX))
  const top = o.y + r.height > window.innerHeight - EDGE_PX ? Math.max(EDGE_PX, o.y - r.height) : o.y
  pos.value = { left: `${left}px`, top: `${top}px` }
})

const size = computed(() => (menu.dialog.value?.kind === 'size' ? menu.dialog.value : null))
const doomed = computed(() => (menu.dialog.value?.kind === 'deleteSheet' ? menu.dialog.value.sheet : null))
const hidden = computed(() => (menu.dialog.value?.kind === 'unhide' ? menu.dialog.value.sheets : null))
const grouping = computed(() => (menu.dialog.value?.kind === 'outline' ? menu.dialog.value : null))
const input = ref<HTMLInputElement | null>(null)
watch(size, async (d, was) => {
  if (!d || was) return
  await nextTick()
  input.value?.focus()
  input.value?.select()
})
const firstChoice = ref<HTMLInputElement | null>(null)
watch(grouping, async (d, was) => {
  if (!d || was) return
  await nextTick()
  firstChoice.value?.focus()
})
const closeDialog = () => (menu.dialog.value = null)

function onDocDown(e: PointerEvent): void {
  if (menu.open.value && !box.value?.contains(e.target as Node)) menu.close()
}
function onDocKey(e: KeyboardEvent): void {
  if (e.key !== 'Escape' || (!menu.open.value && !menu.dialog.value)) return
  e.preventDefault()
  menu.close()
  closeDialog()
}
onMounted(() => {
  document.addEventListener('pointerdown', onDocDown, true)
  document.addEventListener('keydown', onDocKey)
})
onBeforeUnmount(() => {
  document.removeEventListener('pointerdown', onDocDown, true)
  document.removeEventListener('keydown', onDocKey)
})
</script>

<template>
  <div v-if="menu.open.value" ref="box" class="xs-menu" role="menu" :aria-label="menu.open.value.kind === 'cells' ? m.cellMenu : m.menu" :style="pos" data-xs-menu-box>
    <div class="xs-menu-caption">{{ menu.caption.value }}</div>
    <template v-for="item in menu.items.value" :key="item.id">
      <div v-if="item.group" class="xs-menu-sep" role="separator"></div>
      <button type="button" role="menuitem" class="xs-menu-item" :data-xs-menu="item.id" :disabled="item.disabled" @click="menu.choose(item)">{{ item.label }}</button>
    </template>
  </div>
  <div v-if="menu.dialog.value" class="xs-dialog-back" @pointerdown.self="closeDialog">
    <div class="xs-dialog" role="dialog" aria-modal="true" data-xs-dialog>
      <template v-if="size">
        <label class="xs-dialog-field">
          <span class="xs-dialog-title">{{ size.axis === 'columns' ? m.columnWidth : m.rowHeight }}</span>
          <input ref="input" v-model="size.draft" class="xs-dialog-input" type="text" inputmode="decimal" autocomplete="off" spellcheck="false" data-xs-dialog-input @keydown.enter.prevent="menu.applySize()" />
        </label>
        <p v-if="size.error" class="xs-dialog-error" role="alert">{{ size.error }}</p>
      </template>
      <template v-else-if="doomed !== null">
        <p class="xs-dialog-title">{{ m.deleteSheet }}</p>
        <p class="xs-dialog-text">{{ m.deleteSheetText(doomed) }}</p>
      </template>
      <template v-else-if="grouping">
        <p class="xs-dialog-title">{{ grouping.op === 'group' ? m.menuGroup : m.menuUngroup }}</p>
        <div class="xs-dialog-choices" role="radiogroup" @keydown.enter.prevent="menu.applyOutline()">
          <label class="xs-dialog-choice"><input ref="firstChoice" v-model="grouping.axis" type="radio" value="rows" data-xs-outline-axis="rows" />{{ m.outlineRows }}</label>
          <label class="xs-dialog-choice"><input v-model="grouping.axis" type="radio" value="columns" data-xs-outline-axis="columns" />{{ m.outlineColumns }}</label>
        </div>
      </template>
      <template v-else-if="hidden">
        <p class="xs-dialog-title">{{ m.unhideSheet }}</p>
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
    </div>
  </div>
</template>

<style src="../menu.css"></style>
