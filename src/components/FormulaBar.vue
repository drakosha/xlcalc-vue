<script setup lang="ts">
// The bar over the grid, as in Excel: back, the address box, the formula with clickable references (the edit field too), link and lock.
import { computed } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import { safeHref } from '../core/cell-view'
import { useSheetContext } from './context'
import { IconBack, IconFilter, IconLink, IconLock, IconLockOpen, IconSearch } from './icons'

const ctx = useSheetContext()
const { sel, f, nav, ed, v } = ctx
const m = ctx.messages

const editingBar = computed(() => ctx.editor.value?.place === 'bar')
const editingCell = computed(() => ctx.editor.value?.place === 'cell')
const active = computed(() => ctx.selected.value)
const activeLink = computed(() => {
  const a = active.value
  return a ? ctx.w.cellData(a.row, a.col, cellAddress(a.row, a.col)).link : undefined
})
/** What the bar shows: the draft of a field open in the cell, else the text the field opens with, as Excel's formula bar. */
const barText = computed(() => {
  if (editingCell.value) return ctx.editor.value!.text
  const a = active.value
  return a ? ed.editText(a.row, a.col) : ''
})
/** The window's formula of the active cell, when that is the bar's text, shows in pieces with clickable references. */
const showParts = computed(() => !editingCell.value && f.formulaParts.value.length > 0 && barText.value === `=${f.selectedFormula.value}`)
/** With a range selected, the text is the active cell's: it is named. */
const activeName = computed(() => (sel.multi.value && active.value && barText.value ? `${cellAddress(active.value.row, active.value.col)}: ` : ''))

/** A press on the bar keeps the focus where it is (an open field must not blur and commit); the click opens the field here. */
function onBarDown(e: MouseEvent): void {
  if (!ctx.editable.value || editingBar.value || (e.target as HTMLElement).closest('button,a')) return
  e.preventDefault()
}
function onBarClick(e: MouseEvent): void {
  if (!ctx.editable.value || editingBar.value || (e.target as HTMLElement).closest('button,a')) return
  const open = ctx.editor.value
  if (open) {
    // A field open in the cell moves to the bar with its draft.
    open.place = 'bar'
    open.caret = open.text.length
    ed.focusEditor()
    return
  }
  ed.startEdit({ mode: 'edit', place: 'bar' })
}
/** A finger has no funnels on the letters (too small to hit): one in the bar opens the active cell's column. */
const barFunnel = computed(() => v.touch.value && ctx.filter.available.value && ctx.dataRows.value > 1)
function openBarFunnel(e: MouseEvent): void {
  const col = active.value?.col ?? 1
  if (ctx.filter.open.value === col) return ctx.filter.close()
  void ctx.filter.openFunnel(col, (e.currentTarget as HTMLElement).getBoundingClientRect())
}
const setBarInput = (el: unknown) => (ctx.barInput.value = (el as HTMLTextAreaElement | null) ?? null)

function onAddressFocus(e: FocusEvent): void {
  nav.addressFocused.value = true
  ;(e.target as HTMLInputElement).select()
}
function onAddressBlur(): void {
  nav.addressFocused.value = false
  nav.addressDraft.value = sel.selectionRef.value
}
async function onAddressEnter(e: KeyboardEvent): Promise<void> {
  const input = e.target as HTMLInputElement
  await nav.goAddress()
  input.blur()
}
function onAddressEsc(e: KeyboardEvent): void {
  nav.addressDraft.value = sel.selectionRef.value
  ;(e.target as HTMLInputElement).blur()
}
</script>

<template>
  <div class="xs-bar" :class="{ 'xs-bar-editing': editingBar }">
    <button
      type="button"
      class="xs-tool"
      :disabled="!nav.history.value.length"
      :title="nav.history.value.length ? m.back : m.backEmpty"
      @click="nav.goBack()"
    >
      <IconBack />
    </button>
    <input
      v-model="nav.addressDraft.value"
      class="xs-address"
      type="text"
      spellcheck="false"
      placeholder="A1"
      :title="m.address"
      @focus="onAddressFocus"
      @blur="onAddressBlur"
      @keydown.enter.prevent="onAddressEnter"
      @keydown.esc.prevent="onAddressEsc"
    />
    <div class="xs-fx" :class="{ 'xs-fx-edit': ctx.editable.value }" :title="editingBar ? undefined : barText || (ctx.editable.value ? m.editHint : '')" @mousedown="onBarDown" @click="onBarClick">
      <span v-if="!editingBar" class="xs-fx-text">
        <span v-if="activeName" class="xs-muted">{{ activeName }}</span>
        <template v-if="showParts">
          <span>=</span>
          <template v-for="(part, i) in f.formulaParts.value" :key="i">
            <button
              v-if="part.ref && part.color && f.refSheet(part.ref) >= 0"
              type="button"
              class="xs-ref"
              :style="{ color: part.color }"
              :title="part.ref.sheet ? m.goSheet(part.ref.sheet) : m.showRange"
              @click="nav.goRef(part.ref)"
            >
              {{ part.text }}
            </button>
            <span v-else>{{ part.text }}</span>
          </template>
        </template>
        <template v-else>{{ barText }}</template>
      </span>
      <textarea
        v-if="ctx.editable.value"
        :ref="setBarInput"
        :class="['xs-fx-input', editingBar ? '' : 'xs-fx-idle']"
        :value="editingBar ? ctx.editor.value!.text : ''"
        rows="1"
        spellcheck="false"
        autocomplete="off"
        autocapitalize="off"
        :tabindex="editingBar ? 0 : -1"
        :aria-hidden="editingBar ? undefined : 'true'"
        @input="ed.onEditorInput"
        @keydown="ed.onEditorKey"
        @blur="ed.onEditorBlur($event, 'bar')"
      ></textarea>
    </div>
    <template v-if="activeLink && !editingBar">
      <a
        v-if="safeHref(activeLink.href)"
        class="xs-bar-link"
        :href="activeLink.href"
        target="_blank"
        rel="noopener noreferrer nofollow"
        :title="activeLink.href"
        @click="ctx.onLink($event, activeLink, active!.row, active!.col)"
        >{{ activeLink.href }}</a
      >
      <span v-else class="xs-bar-link xs-muted" :title="activeLink.location ?? activeLink.href ?? activeLink.target">{{
        activeLink.location ? m.linkInBook(activeLink.location) : (activeLink.href ?? activeLink.target)
      }}</span>
    </template>
    <div class="xs-tools">
      <button v-if="ctx.props.cellUrl && sel.selectionRef.value" type="button" class="xs-tool" :title="m.cellLink" @click="nav.copyCellLink()">
        <IconLink />
      </button>
      <button
        v-if="v.bookFreeze.value"
        type="button"
        :class="['xs-tool', v.freezeOn.value ? 'xs-tool-on' : '']"
        :title="m.freeze(ctx.sheet.value?.freeze?.rows ?? 0, ctx.sheet.value?.freeze?.cols ?? 0, v.freezeOn.value)"
        @click="v.toggleFreeze()"
      >
        <IconLock v-if="v.freezeOn.value" />
        <IconLockOpen v-else />
      </button>
      <button
        v-if="ctx.find.available.value"
        type="button"
        :class="['xs-tool', ctx.find.isOpen.value ? 'xs-tool-on' : '']"
        :title="m.find"
        data-xs-find
        @click="ctx.find.isOpen.value ? ctx.find.closeFind() : ctx.find.openFind()"
      >
        <IconSearch />
      </button>
      <button
        v-if="barFunnel"
        type="button"
        :class="['xs-tool', 'xs-funnel-tool', ctx.filter.active.value ? 'xs-tool-on' : '']"
        :title="m.filter"
        data-xs-funnel
        @click="openBarFunnel"
      >
        <IconFilter />
      </button>
    </div>
  </div>
</template>
