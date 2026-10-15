<script setup lang="ts">
// The bar over the grid, as in Excel: back, the address box, the formula's field (its references coloured, its list of functions and tip of arguments) one line high or expanded, link and lock.
import { computed, onBeforeUnmount, shallowRef, watch } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { GridReference } from 'xlcalc/contract'
import { safeHref } from '../core/cell-view'
import { onMac } from '../core/clip'
import type { FormulaPart } from '../core/formula'
import { canPaint, paintRefs, type Painted } from '../core/ref-paint'
import type { ValidationRect } from '../core/validation'
import { useSheetContext } from './context'
import FunctionList from './FunctionList.vue'
import FunctionTip from './FunctionTip.vue'
import { IconBack, IconDown, IconFilter, IconLink, IconLock, IconLockOpen, IconSearch, IconUp } from './icons'

const ctx = useSheetContext()
const { sel, f, nav, ed, v } = ctx
const m = ctx.messages

const editingBar = computed(() => ctx.editor.value?.place === 'bar')
const active = computed(() => ctx.selected.value)
const activeLink = computed(() => {
  const a = active.value
  return a ? ctx.w.cellData(a.row, a.col, cellAddress(a.row, a.col)).link : undefined
})
/** What the field holds: the draft of an open field, in the cell too, else the text the active cell opens with, as Excel's formula bar. */
const barText = computed(() => {
  const open = ctx.editor.value
  if (open) return open.text
  const a = active.value
  return a ? ed.editText(a.row, a.col) : ''
})
/** Outside an edit the field takes no typing where none can start: a read-only sheet, a locked cell, no active cell. */
const barReadonly = computed(() => {
  const a = active.value
  return !ctx.editor.value && (!ctx.editable.value || !a || ed.lockAt(a.row, a.col) !== null)
})
/** The text in pieces for the layer over the field: a formula's references in the colours of their frames on the grid (the draft's while one is typed), else the text whole. */
const parts = computed<FormulaPart[]>(() => {
  const text = barText.value
  const refs = f.formulaParts.value
  return refs.length && text === `=${f.shownFormula.value}` ? [{ text: '=' }, ...refs] : [{ text }]
})
/** The layer's text: a last line break shows an empty line in the field, so the layer needs a character on it to have that line too. */
const layerText = computed(() => (barText.value.endsWith('\n') ? `${barText.value} ` : barText.value))
/** Highlights colour the references on the layer's one text node, which then lies on the field's text to the pixel; without them, spans. */
const painting = canPaint()
let painted: Painted = []

/** The reference whose characters the caret stands at (either edge counts), when its sheet is in the book. */
function refAt(at: number): GridReference | null {
  let from = 0
  for (const part of parts.value) {
    const to = from + part.text.length
    if (part.ref && at >= from && at <= to && f.refSheet(part.ref) >= 0) return part.ref
    from = to
  }
  return null
}
/** Outside an edit with the mouse, Ctrl+click (Cmd+click) goes to a reference, as in a code editor: said in the field's tip. */
const goTip = computed(() => (!ctx.editor.value && !v.touch.value && parts.value.some((p) => p.ref && f.refSheet(p.ref) >= 0) ? m.value.goRefHint(onMac() ? 'Cmd' : 'Ctrl') : undefined))
/** A press with Ctrl or Cmd outside an edit: its click goes to the reference under the caret instead of editing. */
let going = false
function onFieldDown(e: MouseEvent): void {
  going = !ctx.editor.value && !v.touch.value && (e.ctrlKey || e.metaKey)
}
/** The field's focus starts the edit in the bar where the press put the caret or the selection, as F2 does; a field open in the cell moves here with its draft. */
function onFieldFocus(e: FocusEvent): void {
  if (going || !ctx.editable.value) return
  const open = ctx.editor.value
  if (open) return void (open.place = 'bar')
  const el = e.target as HTMLTextAreaElement
  ed.startEdit({ mode: 'edit', place: 'bar', caret: el.selectionStart, caretEnd: el.selectionEnd })
}
function onFieldClick(e: MouseEvent): void {
  const el = e.target as HTMLTextAreaElement
  if (going) {
    going = false
    const ref = refAt(el.selectionStart)
    if (ref) {
      ctx.scroller.value?.focus({ preventScroll: true })
      return void nav.goRef(ref)
    }
  }
  if (ctx.editor.value) return ed.onEditorClick()
  if (!barReadonly.value) ed.startEdit({ mode: 'edit', place: 'bar', caret: el.selectionStart, caretEnd: el.selectionEnd })
}
function onFieldBlur(e: FocusEvent): void {
  going = false
  ed.onEditorBlur(e, 'bar')
}
/** A key in the field outside an edit (the edit ended while it kept the focus) starts one, so the field never parts from the cell; on a read-only field Esc goes back to the grid. */
function onFieldKey(e: KeyboardEvent): void {
  if (!ctx.editor.value && !e.isComposing && !['Shift', 'Control', 'Alt', 'Meta'].includes(e.key)) {
    const el = e.target as HTMLTextAreaElement
    if (!barReadonly.value) ed.startEdit({ mode: 'edit', place: 'bar', caret: el.selectionStart, caretEnd: el.selectionEnd })
    else if (e.key === 'Escape') ctx.scroller.value?.focus({ preventScroll: true })
  }
  ed.onEditorKey(e)
}
/** The layer scrolls with the field; outside an edit in the bar a new text shows from its top. */
const layer = shallowRef<HTMLElement | null>(null)
function syncLayer(): void {
  const el = ctx.barInput.value
  if (el && layer.value) layer.value.scrollTop = el.scrollTop
}
watch(
  barText,
  () => {
    const el = ctx.barInput.value
    if (el && !editingBar.value) el.scrollTop = 0
    syncLayer()
  },
  { flush: 'post' },
)
if (painting) {
  watch(
    [parts, layerText],
    () => {
      const pieces = []
      let from = 0
      for (const part of parts.value) {
        if (part.color) pieces.push({ from, to: from + part.text.length, color: part.color })
        from += part.text.length
      }
      const node = layer.value?.firstChild
      painted = paintRefs(node instanceof Text ? node : null, pieces, painted)
    },
    { flush: 'post', immediate: true },
  )
  onBeforeUnmount(() => (painted = paintRefs(null, [], painted)))
}
/** A finger has no funnels on the letters (too small to hit): one in the bar opens the active cell's column. */
const barFunnel = computed(() => v.touch.value && ctx.filter.available.value && ctx.dataRows.value > 1)
function openBarFunnel(e: MouseEvent): void {
  const col = active.value?.col ?? 1
  if (ctx.filter.open.value === col) return ctx.filter.close()
  void ctx.filter.openFunnel(col, (e.currentTarget as HTMLElement).getBoundingClientRect())
}
const setBarInput = (el: unknown) => (ctx.barInput.value = (el as HTMLTextAreaElement | null) ?? null)

/** The list of functions or the tip of arguments under the bar's field, in the formula's coordinates, kept inside the window: measured as one opens, as the text changes and as the bar expands or collapses. */
const listPlace = shallowRef<{ rect: ValidationRect; view: ValidationRect } | null>(null)
watch(
  () => [editingBar.value && (ed.ac.open.value || !!ed.tip.parts.value), ctx.editor.value?.text, ctx.barOpen.value] as const,
  ([shown]) => {
    const input = ctx.barInput.value
    const fx = input?.parentElement
    if (!shown || !input || !fx) return void (listPlace.value = null)
    const f = fx.getBoundingClientRect()
    const i = input.getBoundingClientRect()
    const gutter = 8
    listPlace.value = {
      rect: { top: i.top - f.top, left: i.left - f.left, width: i.width, height: i.height },
      view: { top: -f.top, left: gutter - f.left, width: document.documentElement.clientWidth - 2 * gutter, height: window.innerHeight },
    }
  },
  { flush: 'post' },
)

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
    <div class="xs-fx-line">
      <div class="xs-fx" :class="{ 'xs-fx-open': ctx.barOpen.value }">
        <textarea
          :ref="setBarInput"
          class="xs-fx-input"
          :value="barText"
          :readonly="barReadonly"
          :title="goTip"
          rows="1"
          spellcheck="false"
          autocomplete="off"
          autocapitalize="off"
          @mousedown="onFieldDown"
          @focus="onFieldFocus"
          @click="onFieldClick"
          @input="ed.onEditorInput"
          @keydown="onFieldKey"
          @scroll="syncLayer"
          @blur="onFieldBlur"
        ></textarea>
        <!-- the same text over the transparent field's, its references coloured; the caret and the selection are the field's -->
        <div v-if="painting" ref="layer" class="xs-fx-layer" aria-hidden="true">{{ layerText }}</div>
        <div v-else ref="layer" class="xs-fx-layer" aria-hidden="true"><template v-for="(part, i) in parts" :key="i"><span v-if="part.color" class="xs-ref" :style="{ color: part.color }">{{ part.text }}</span><template v-else>{{ part.text }}</template></template>{{ layerText.slice(barText.length) }}</div>
        <FunctionList
          v-if="listPlace && editingBar && ed.ac.open.value"
          :rect="listPlace.rect"
          :view="listPlace.view"
          :items="ed.ac.items.value"
          :active="ed.ac.index.value"
          :touch="ctx.v.touch.value"
          :messages="m"
          @highlight="(i) => (ed.ac.index.value = i)"
          @pick="(i) => ed.ac.accept(i)"
        />
        <FunctionTip v-if="listPlace && editingBar && ed.tip.parts.value" :rect="listPlace.rect" :view="listPlace.view" :parts="ed.tip.parts.value" />
      </div>
      <!-- Excel's arrow at the bar's right end; a press keeps an open field open -->
      <button
        type="button"
        class="xs-tool xs-fx-toggle"
        data-xs-bar-toggle
        :title="ctx.barOpen.value ? m.collapseBar : m.expandBar"
        :aria-label="ctx.barOpen.value ? m.collapseBar : m.expandBar"
        :aria-expanded="ctx.barOpen.value ? 'true' : 'false'"
        aria-keyshortcuts="Control+Shift+U"
        @mousedown.prevent
        @click="ctx.toggleBar()"
      >
        <IconUp v-if="ctx.barOpen.value" />
        <IconDown v-else />
      </button>
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
