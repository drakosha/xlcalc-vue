<script setup lang="ts">
// The bar above the grid, as in Excel: Back, the Name Box, the formula input (with colored references, the function list and the argument tooltip), one line high or expanded, the link and the lock.
import { computed, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { GridReference } from 'xlcalc/contract'
import { safeHref } from '../core/cell-view'
import { isMod, onMac } from '../core/keys'
import type { FormulaPart } from '../core/formula'
import { canPaint, paintRefs, pieceAtPoint, type Painted } from '../core/ref-paint'
import type { PopupRect } from './popup'
import { useSheetContext } from './context'
import { refInto } from './dom'
import FunctionList from './FunctionList.vue'
import FunctionTip from './FunctionTip.vue'
import InsertFunction from './InsertFunction.vue'
import { IconBack, IconCheck, IconClose, IconDown, IconFilter, IconFx, IconLink, IconLock, IconLockOpen, IconSearch, IconUp } from './icons'

const ctx = useSheetContext()
const { sel, f, nav, ed, v } = ctx
const m = ctx.messages

const editingBar = computed(() => ctx.editor.value?.place === 'bar')
const active = computed(() => ctx.selected.value)
const activeLink = computed(() => {
  const a = active.value
  return a ? ctx.w.cellData(a.row, a.col, cellAddress(a.row, a.col)).link : undefined
})
/** What the input contains: the draft of an open editor, in the cell too, otherwise the text the active cell opens with, as in Excel's formula bar. */
const barText = computed(() => {
  const open = ctx.editor.value
  if (open) return open.text
  const a = active.value
  return a ? ed.editText(a.row, a.col) : ''
})
/** Outside editing, the input accepts no typing where no edit can start: a read-only sheet, a locked cell, no active cell. */
const barReadonly = computed(() => {
  const a = active.value
  return !ctx.editor.value && (!ctx.editable.value || !a || ed.lockAt(a.row, a.col) !== null)
})
/** The text in pieces for the layer over the input: a formula's references in the colors of their range borders on the grid (the draft's while one is typed), otherwise the text as a whole. */
const parts = computed<FormulaPart[]>(() => {
  const text = barText.value
  const refs = f.formulaParts.value
  return refs.length && text === `=${f.shownFormula.value}` ? [{ text: '=' }, ...refs] : [{ text }]
})
/** The layer's text: a final line break shows an empty line in the input, so the layer needs a character on it to have that line too. */
const layerText = computed(() => (barText.value.endsWith('\n') ? `${barText.value} ` : barText.value))
/** Highlights color the references on the layer's single text node, which then lines up with the input's text to the pixel; without the API, spans. */
const painting = canPaint()
let painted: Painted = []

/** The references one can go to (their sheet is in the workbook): their text and offsets. */
const goable = computed(() => {
  const out: { part: number; from: number; to: number; ref: GridReference }[] = []
  let from = 0
  parts.value.forEach((part, i) => {
    const to = from + part.text.length
    if (part.ref && f.refSheet(part.ref) >= 0) out.push({ part: i, from, to, ref: part.ref })
    from = to
  })
  return out
})
/** The reference whose characters the caret is at (either edge counts). */
const refAt = (at: number): GridReference | null => goable.value.find((g) => at >= g.from && at <= g.to)?.ref ?? null
/** Outside editing with a mouse, Ctrl+click (Cmd+click) goes to a reference, as in a code editor: the input's tooltip says so. */
const goTip = computed(() => (!ctx.editor.value && !v.touch.value && goable.value.length ? m.value.goRefHint(onMac() ? 'Cmd' : 'Ctrl') : undefined))
/** A press with Ctrl or Cmd outside editing: its click goes to the reference under the caret instead of editing. */
let going = false
function onFieldDown(e: MouseEvent): void {
  // a mouse press a browser may send after a finger's long press that already navigated: the input takes no focus
  if (ctx.menu.press.after()) return void e.preventDefault()
  going = !ctx.editor.value && !v.touch.value && isMod(e)
}
/** Focusing the input starts editing in the formula bar where the press put the caret or the selection, as F2 does; an editor open in the cell moves here with its draft. */
function onFieldFocus(e: FocusEvent): void {
  if (going || ctx.menu.press.after() || !ctx.editable.value) return
  const open = ctx.editor.value
  if (open) return void (open.place = 'bar')
  const el = e.target as HTMLTextAreaElement
  ed.startEdit({ mode: 'edit', place: 'bar', caret: el.selectionStart, caretEnd: el.selectionEnd })
}
function onFieldClick(e: MouseEvent): void {
  const el = e.target as HTMLTextAreaElement
  if (ctx.menu.press.after()) return
  if (going) {
    going = false
    const ref = goable.value[hovered.value]?.ref ?? refAt(el.selectionStart)
    if (ref) return goRef(ref)
  }
  if (ctx.editor.value) return ed.onEditorClick()
  if (!barReadonly.value) ed.startEdit({ mode: 'edit', place: 'bar', caret: el.selectionStart, caretEnd: el.selectionEnd })
}
function onFieldBlur(e: FocusEvent): void {
  going = false
  ed.onEditorBlur(e, 'bar')
}
/** A key in the input outside editing (the edit ended while it kept focus) starts one, so the input always follows the cell; on a read-only input Esc returns to the grid. */
function onFieldKey(e: KeyboardEvent): void {
  if (!ctx.editor.value && !e.isComposing && !['Shift', 'Control', 'Alt', 'Meta'].includes(e.key)) {
    const el = e.target as HTMLTextAreaElement
    if (!barReadonly.value) ed.startEdit({ mode: 'edit', place: 'bar', caret: el.selectionStart, caretEnd: el.selectionEnd })
    else if (e.key === 'Escape') ctx.focusGrid()
  }
  ed.onEditorKey(e)
}
/** Goes to a reference's range from the formula bar: the grid takes focus, as after the Name Box. */
function goRef(ref: GridReference): void {
  ctx.focusGrid()
  nav.goRef(ref)
}
/** The layer scrolls with the input; outside editing in the formula bar a new text is shown from its top. */
const layer = shallowRef<HTMLElement | null>(null)

/** While Ctrl or Cmd is held outside editing, the reference under the pointer is underlined and the pointer becomes a hand, like a link in a code editor; either may come first. */
const modHeld = ref(false)
let mouseAt: { x: number; y: number } | null = null
/** The index in `goable` of the reference under the pointer while it would navigate, otherwise -1. */
const hovered = ref(-1)
function updateHover(): void {
  const on = modHeld.value && mouseAt && layer.value && !ctx.editor.value && !v.touch.value
  hovered.value = on ? pieceAtPoint(layer.value!, goable.value, mouseAt!.x, mouseAt!.y) : -1
}
const hoveredPart = computed(() => goable.value[hovered.value]?.part ?? -1)
function onFieldMove(e: MouseEvent): void {
  modHeld.value = isMod(e)
  mouseAt = { x: e.clientX, y: e.clientY }
  updateHover()
}
function onFieldLeave(): void {
  mouseAt = null
  updateHover()
}
function onModKey(e: KeyboardEvent): void {
  if (e.key !== 'Control' && e.key !== 'Meta') return
  modHeld.value = isMod(e)
  updateHover()
}
function onWindowBlur(): void {
  modHeld.value = false
  updateHover()
}
watch([goable, () => ctx.editor.value, v.touch], updateHover, { flush: 'post' })
onMounted(() => {
  window.addEventListener('keydown', onModKey, true)
  window.addEventListener('keyup', onModKey, true)
  window.addEventListener('blur', onWindowBlur)
})
onBeforeUnmount(() => {
  window.removeEventListener('keydown', onModKey, true)
  window.removeEventListener('keyup', onModKey, true)
  window.removeEventListener('blur', onWindowBlur)
})

// A long press on a reference outside editing goes to it, with the timing and slop of the sheet's long presses; the browser's menu and text selection for that press are suppressed.
/** How far beside a reference's characters a finger still lands on it, in px; how far the finger may then move is the long press's own slop. */
const FINGER_SLOP_PX = 6
let pressRef: GridReference | null = null
const pressHeld = ref(false)
function onFieldPointerDown(e: PointerEvent): void {
  pressRef = null
  if (e.pointerType !== 'touch' || ctx.editor.value || !layer.value) return
  const i = pieceAtPoint(layer.value, goable.value, e.clientX, e.clientY, FINGER_SLOP_PX)
  if (i < 0) return
  const ref = (pressRef = goable.value[i]!.ref)
  pressHeld.value = true
  ctx.menu.press.start(e, () => {
    pressHeld.value = false
    goRef(ref)
  })
}
function onFieldPointerMove(e: PointerEvent): void {
  ctx.menu.press.move(e)
  if (!ctx.menu.press.pressing()) pressHeld.value = false
}
function onFieldPointerEnd(): void {
  ctx.menu.press.cancel()
  pressHeld.value = false
}
/** A browser's long press may bring up its menu before the press's time: the press is canceled then; the menu that follows it is suppressed too. */
function onFieldContext(e: MouseEvent): void {
  const pending = pressHeld.value && pressRef
  if (!pending && !ctx.menu.press.after()) return
  e.preventDefault()
  if (!pending) return
  onFieldPointerEnd()
  goRef(pending)
}
function onFieldSelectStart(e: Event): void {
  if (pressHeld.value || ctx.menu.press.after()) e.preventDefault()
}
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
    [parts, layerText, hoveredPart],
    () => {
      const pieces = []
      let from = 0
      for (const [i, part] of parts.value.entries()) {
        if (part.color) pieces.push({ from, to: from + part.text.length, color: part.color, underline: i === hoveredPart.value })
        from += part.text.length
      }
      const node = layer.value?.firstChild
      painted = paintRefs(node instanceof Text ? node : null, pieces, painted)
    },
    { flush: 'post', immediate: true },
  )
  onBeforeUnmount(() => (painted = paintRefs(null, [], painted)))
}
/** Excel's Insert Function: on an editable sheet whose source lists its functions; its dialog keeps an open edit open. */
const canInsert = computed(() => ctx.editable.value && !!ctx.source().functionNames)
const inserting = ref(false)
function onInsertPick(name: string): void {
  inserting.value = false
  ed.insertFunction(name)
}
function onInsertClose(): void {
  inserting.value = false
  if (ctx.editor.value) ed.focusEditor()
  else ctx.focusGrid()
}
/** Touch input gets no funnels on the column letters (too small to hit): one in the formula bar opens the active cell's column. */
const barFunnel = computed(() => v.touch.value && ctx.filter.available.value && ctx.dataRows.value > 1)
function openBarFunnel(e: MouseEvent): void {
  const col = active.value?.col ?? 1
  if (ctx.filter.open.value === col) return ctx.filter.close()
  void ctx.filter.openFunnel(col, (e.currentTarget as HTMLElement).getBoundingClientRect())
}
const setBarInput = refInto(ctx.barInput)

/** The function list or the argument tooltip below the formula bar's input, in the formula's coordinates, kept inside the viewport: measured as one opens, as the text changes and as the bar expands or collapses. */
const listPlace = shallowRef<{ rect: PopupRect; view: PopupRect } | null>(null)
watch(
  () => [editingBar.value && (ed.ac.open.value || !!ed.tip.parts.value), ctx.editor.value?.text, ctx.barOpen.value] as const,
  ([shown]) => {
    const input = ctx.barInput.value
    const fx = input?.parentElement
    if (!shown || !input || !fx) return void (listPlace.value = null)
    const box = fx.getBoundingClientRect()
    const i = input.getBoundingClientRect()
    const gutter = 8
    listPlace.value = {
      rect: { top: i.top - box.top, left: i.left - box.left, width: i.width, height: i.height },
      view: { top: -box.top, left: gutter - box.left, width: document.documentElement.clientWidth - 2 * gutter, height: window.innerHeight },
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
      <!-- Excel's Cancel, Enter and Insert Function to the left of the input; a press keeps an open editor's focus until the button acts -->
      <button type="button" class="xs-tool" :disabled="!ctx.editor.value" :title="m.cancel" :aria-label="m.cancel" data-xs-bar-cancel @mousedown.prevent @click="ed.cancel()">
        <IconClose />
      </button>
      <button type="button" class="xs-tool" :disabled="!ctx.editor.value" :title="m.barEnter" :aria-label="m.barEnter" data-xs-bar-enter @mousedown.prevent @click="ed.commit(0, 0)">
        <IconCheck />
      </button>
      <button type="button" class="xs-tool" :disabled="!canInsert" :title="m.insertFunction" :aria-label="m.insertFunction" data-xs-bar-fx @mousedown.prevent @click="inserting = true">
        <IconFx />
      </button>
      <div class="xs-fx" :class="{ 'xs-fx-open': ctx.barOpen.value }">
        <textarea
          :ref="setBarInput"
          :class="['xs-fx-input', { 'xs-fx-go': hovered >= 0, 'xs-fx-held': pressHeld }]"
          :value="barText"
          :readonly="barReadonly"
          :title="goTip"
          rows="1"
          spellcheck="false"
          autocomplete="off"
          autocapitalize="off"
          @mousedown="onFieldDown"
          @mousemove="onFieldMove"
          @mouseleave="onFieldLeave"
          @pointerdown="onFieldPointerDown"
          @pointermove="onFieldPointerMove"
          @pointerup="onFieldPointerEnd"
          @pointercancel="onFieldPointerEnd"
          @contextmenu="onFieldContext"
          @selectstart="onFieldSelectStart"
          @focus="onFieldFocus"
          @click="onFieldClick"
          @input="ed.onEditorInput"
          @keydown="onFieldKey"
          @scroll="syncLayer"
          @blur="onFieldBlur"
        ></textarea>
      <!-- the same text over the transparent input's, with its references colored; the caret and the selection are the input's -->
        <div v-if="painting" ref="layer" class="xs-fx-layer" aria-hidden="true">{{ layerText }}</div>
        <div v-else ref="layer" class="xs-fx-layer" aria-hidden="true"><template v-for="(part, i) in parts" :key="i"><span v-if="part.color" :class="['xs-ref', { 'xs-ref-go': i === hoveredPart }]" :style="{ color: part.color }">{{ part.text }}</span><template v-else>{{ part.text }}</template></template>{{ layerText.slice(barText.length) }}</div>
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
      <!-- Excel's arrow at the bar's right end; a press keeps an open editor open -->
      <button
        type="button"
        class="xs-tool"
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
    <InsertFunction v-if="inserting" @pick="onInsertPick" @close="onInsertClose" />
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
