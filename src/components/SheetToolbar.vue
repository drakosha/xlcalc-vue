<script setup lang="ts">
// Excel's formatting toolbar over a sheet source: undo and redo, bold, italic, underline, font and fill colours, borders,
// horizontal alignment, wrapping, Merge & Center with its list (through the sheet, with the source's `structure`) and
// Excel's Number Format box (xlcalc's, in the book's locale) over the selection the sheet's `state` event gives, and
// Format Cells for a host that shows the dialog. On a narrow frame what does not fit
// goes into a menu; on a finger every target is 44 × 44 px.
import { computed, getCurrentInstance, onBeforeUnmount, onMounted, reactive, ref, watch } from 'vue'
import type { SheetViewApi } from '../core/commands'
import type { MergeCommand } from '../core/merge'
import type { SheetViewState } from '../core/sheet'
import { useMessages, type SheetMessages } from '../messages'
import { useNumberFormats } from '../number-formats'
import type { EditOutcome, FormatPatch } from '../source'
import {
  alignToggle,
  borderPatches,
  FIRST_FILL_COLOUR,
  FIRST_FONT_COLOUR,
  fontOn,
  fontToggle,
  upperColour,
  wrapToggle,
  type BorderKind,
  type FontToggle,
  type HorizontalAlign,
  type ToolbarSource,
} from '../toolbar'
import ColourSwatches from './ColourSwatches.vue'
import { IconDown } from './icons'
import { useOutsidePress } from './popup'
import {
  IconAlignCenter,
  IconAlignLeft,
  IconAlignRight,
  IconBold,
  IconBorderAll,
  IconBorderBottom,
  IconBorderNone,
  IconBorderOutside,
  IconFillColour,
  IconFontColour,
  IconFormatCells,
  IconItalic,
  IconMerge,
  IconMore,
  IconRedo,
  IconUnderline,
  IconUndo,
  IconWrap,
} from './toolbar-icons'

const props = withDefaults(
  defineProps<{
    source: ToolbarSource
    /** What the sheet's `state` event gave: the selection and the active cell's look; null leaves nothing to press. */
    state: SheetViewState | null
    /** Interface texts; a part overrides the English defaults. */
    messages?: Partial<SheetMessages>
    /** The book's locale (as `SheetView`'s): the Number Format box names its categories and items in it; en-US without one. */
    locale?: string
    /**
     * The sheet (a `SheetView`'s template ref) to format, undo and redo through, as its keys do: it applies the outcome, sets
     * the selection and shows a refusal itself. Without it the toolbar calls the source and emits `select` and `error`.
     */
    sheet?: SheetViewApi | null
  }>(),
  { messages: undefined, locale: undefined, sheet: null },
)
const emit = defineEmits<{
  /** Where the selection goes after an undo or a redo (`EditOutcome.select`), for the sheet's `focus`; not with `sheet`. */
  select: [place: { sheet: string; range: string }]
  /** The source refused or failed a call: its message; not with `sheet`, which shows it. */
  error: [message: string]
  /** A call is done: the host may give the grid its focus back. */
  done: []
  /** Format Cells… pressed, for a host with a `FormatDialog` of its own; without a listener the button opens the sheet's (`sheet`), and without both there is none. */
  formatCells: []
}>()

/** Under this width the toolbar keeps its first buttons in a row and the rest in a menu, as the sheet goes narrow. */
const NARROW_PX = 600

const m = useMessages(() => props.messages)
const instance = getCurrentInstance()
/** Whether the host listens to `formatCells` now, so that the button opens its own dialog instead of the sheet's. */
function hostDialog(): boolean {
  const listener = instance?.vnode.props?.onFormatCells
  return typeof listener === 'function' || Array.isArray(listener)
}
function formatCells(): void {
  if (hostDialog() || !props.sheet) emit('formatCells')
  else props.sheet.formatCells()
}
const style = computed(() => props.state?.style)

const FONT: { what: FontToggle; icon: typeof IconBold; title: keyof SheetMessages }[] = [
  { what: 'bold', icon: IconBold, title: 'bold' },
  { what: 'italic', icon: IconItalic, title: 'italic' },
  { what: 'underline', icon: IconUnderline, title: 'underline' },
]
const ALIGN: { h: HorizontalAlign; icon: typeof IconBold; title: keyof SheetMessages }[] = [
  { h: 'left', icon: IconAlignLeft, title: 'alignLeft' },
  { h: 'center', icon: IconAlignCenter, title: 'alignCenter' },
  { h: 'right', icon: IconAlignRight, title: 'alignRight' },
]
const BORDERS: { kind: BorderKind; icon: typeof IconBold; title: keyof SheetMessages }[] = [
  { kind: 'all', icon: IconBorderAll, title: 'borderAll' },
  { kind: 'outside', icon: IconBorderOutside, title: 'borderOutside' },
  { kind: 'bottom', icon: IconBorderBottom, title: 'borderBottom' },
  { kind: 'none', icon: IconBorderNone, title: 'borderNone' },
]
/** Excel's list under Merge & Center: each item is named by its command's text, its hook (`data-tb`) is the command in kebab case. */
const MERGES: readonly MergeCommand[] = ['mergeCenter', 'mergeAcross', 'mergeCells', 'unmergeCells']
const kebab = (name: string) => name.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)
/** Excel's split buttons of the font's and the fill's colours: the button applies the last colour, the menu the choice without one and the palette. */
type Colouring = 'font' | 'fill'
const COLOURS: { what: Colouring; icon: typeof IconFontColour; title: keyof SheetMessages; none: keyof SheetMessages; noneHook: string }[] = [
  { what: 'font', icon: IconFontColour, title: 'fontColour', none: 'colourAuto', noneHook: 'colour-auto' },
  { what: 'fill', icon: IconFillColour, title: 'fillColour', none: 'fillNone', noneHook: 'fill-none' },
]
const text = (key: keyof SheetMessages) => m.value[key] as string

// ── what the active cell shows ──
/** Excel's box: the list's items, and the name of the category the active cell's format is (General without one). */
const numbers = useNumberFormats(() => props.locale)
const category = computed(() => (props.state ? numbers.categoryOf(props.state.numberFormat) : null))
const fontColour = computed(() => upperColour(style.value?.font?.color))
const fillColour = computed(() => upperColour(style.value?.fill?.color))
/** Bumped after the toolbar's calls and the source's events: `history` is a plain getter, read again (and copied) on them. */
const tick = ref(0)
const history = computed(() => {
  void tick.value
  void props.state
  const h = props.source.history
  return h ? { undo: h.undo, redo: h.redo } : null
})
const canUndo = computed(() => !history.value || history.value.undo > 0)
const canRedo = computed(() => !history.value || history.value.redo > 0)

// ── the frame: narrow, a finger ──
const root = ref<HTMLElement | null>(null)
const width = ref(0)
const coarse = ref(false)
const narrow = computed(() => (width.value || (typeof window === 'undefined' ? NARROW_PX : window.innerWidth)) < NARROW_PX)
const touch = computed(() => narrow.value || coarse.value)
const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(([entry]) => (width.value = entry?.contentRect.width ?? 0))
watch(root, (el, old) => {
  if (old) observer?.unobserve(old)
  if (el) observer?.observe(el)
})
let coarseQuery: MediaQueryList | null = null
const onCoarse = () => (coarse.value = !!coarseQuery?.matches)

// ── menus: the colour palettes, the borders, and the narrow frame's menu ──
const open = ref<'font' | 'fill' | 'borders' | 'merge' | null>(null)
const menu = ref(false)
const toggle = (what: 'font' | 'fill' | 'borders' | 'merge') => (open.value = open.value === what ? null : what)
function closeAll(): void {
  open.value = null
  menu.value = false
}
useOutsidePress(root, closeAll, () => open.value !== null || menu.value)
/** Buttons do not take the focus from the grid, so its keys go on working; the number format list does take it. */
function keepFocus(e: MouseEvent): void {
  if (!(e.target as HTMLElement).closest('select')) e.preventDefault()
}

let unsubscribe: (() => void) | null = null
function subscribe(): void {
  unsubscribe?.()
  unsubscribe = props.source.format ? (props.source.subscribe?.(() => tick.value++) ?? null) : null
}
watch(() => props.source, subscribe)
onMounted(() => {
  subscribe()
  if (typeof window.matchMedia === 'function') {
    coarseQuery = window.matchMedia('(pointer: coarse)')
    onCoarse()
    coarseQuery.addEventListener?.('change', onCoarse)
  }
})
onBeforeUnmount(() => {
  unsubscribe?.()
  observer?.disconnect()
  coarseQuery?.removeEventListener?.('change', onCoarse)
})

// ── calls: through the sheet when there is one, else the source's own ──
async function run(call: () => Promise<EditOutcome | null>): Promise<EditOutcome | null> {
  let out: EditOutcome | null
  try {
    out = await call()
  } catch (e) {
    out = { ok: false, reason: 'failed', message: e instanceof Error ? e.message : '' }
  }
  tick.value++
  if (out && !out.ok && !props.sheet) emit('error', out.message || m.value.editFailed)
  emit('done')
  return out
}
/** Patches over the selection's sheet, in one call and so one step of undo. */
function apply(patches: FormatPatch[]): void {
  const st = props.state
  const format = props.source.format
  const sheet = props.sheet
  if (!st || !format || !patches.length) return
  void run(() => (sheet ? sheet.format(patches) : format.call(props.source, st.sheet, patches)))
}
const over = (p: Omit<FormatPatch, 'range'>): FormatPatch[] => (props.state ? [{ range: props.state.ref, ...p }] : [])

/** Undo and redo one at a time: a second press while one is out sends nothing. */
let stepping = false
async function step(kind: 'undo' | 'redo'): Promise<void> {
  const call = props.source[kind]
  const sheet = props.sheet
  if (!call || stepping) return
  stepping = true
  try {
    const out = await run(() => (sheet ? sheet[kind]() : call.call(props.source)))
    if (out?.ok && out.select && !sheet) emit('select', out.select)
  } finally {
    stepping = false
  }
}

// The split buttons apply the colour picked last, as Excel's do.
const last = reactive<Record<Colouring, string>>({ font: FIRST_FONT_COLOUR, fill: FIRST_FILL_COLOUR })
const shown = computed<Record<Colouring, string | null>>(() => ({ font: fontColour.value, fill: fillColour.value }))
/** A colour, or none (Automatic, No fill), on the text or the fill of the selection. */
function paint(what: Colouring, colour: string | null): void {
  if (colour) last[what] = colour
  open.value = null
  apply(over({ style: what === 'font' ? { font: { color: colour } } : { fill: colour ? { color: colour } : null } }))
}
function border(kind: BorderKind): void {
  open.value = null
  if (props.state) apply(borderPatches(props.state.ref, kind))
}
/** Excel's Merge & Center is pressed on a merged cell, and pressed then it unmerges. */
const merged = computed(() => !!props.state?.merged)
/** A merge command through the sheet, which warns before values are lost and shows a refusal. */
function merge(command: MergeCommand): void {
  open.value = null
  const sheet = props.sheet
  if (sheet && props.state?.canMerge) void run(() => sheet.merge(command))
}
/** An item of the list applies its format; the box goes back to naming the category, as Excel's does. */
function pickNumber(e: Event): void {
  const el = e.target as HTMLSelectElement
  const item = numbers.choices.value.find((x) => x.key === el.value)
  el.value = ''
  if (item) apply(over({ numberFormat: item.format }))
}
</script>

<template>
  <div
    v-if="source.format"
    ref="root"
    :class="['xs-toolbar', { 'xs-tb-narrow': narrow, 'xs-touch': touch }]"
    role="toolbar"
    :aria-label="m.toolbar"
    @mousedown="keepFocus"
    @keydown.esc="closeAll"
  >
    <div class="xs-tb-row">
      <span v-if="source.undo || source.redo" class="xs-tb-group">
        <button v-if="source.undo" type="button" class="xs-tool" data-tb="undo" :title="m.undo" :aria-label="m.undo" :disabled="!canUndo" @click="step('undo')">
          <IconUndo />
        </button>
        <button v-if="source.redo" type="button" class="xs-tool" data-tb="redo" :title="m.redo" :aria-label="m.redo" :disabled="!canRedo" @click="step('redo')">
          <IconRedo />
        </button>
      </span>
      <span class="xs-tb-group">
        <button
          v-for="b in FONT"
          :key="b.what"
          type="button"
          :class="['xs-tool', { 'xs-tool-on': fontOn(style, b.what) }]"
          :data-tb="b.what"
          :title="text(b.title)"
          :aria-label="text(b.title)"
          :aria-pressed="fontOn(style, b.what)"
          :disabled="!state"
          @click="apply(over({ style: fontToggle(style, b.what) }))"
        >
          <component :is="b.icon" />
        </button>
      </span>
      <button v-if="narrow" type="button" :class="['xs-tool', 'xs-tb-more', { 'xs-tool-on': menu }]" data-tb="more" :title="m.toolbarMore" :aria-label="m.toolbarMore" :aria-expanded="menu" @click="menu = !menu">
        <IconMore />
      </button>
    </div>
    <div v-show="!narrow || menu" :class="narrow ? 'xs-tb-menu' : 'xs-tb-row'" data-tb-menu>
      <span v-if="numbers.choices.value.length" class="xs-tb-group">
        <select
          class="xs-field xs-field-s xs-tb-select"
          data-tb="number"
          :data-tb-category="category?.category"
          :title="m.numberFormat"
          :aria-label="m.numberFormat"
          value=""
          :disabled="!state"
          @change="pickNumber"
        >
          <option value="" hidden>{{ category?.name ?? '' }}</option>
          <option v-for="c in numbers.choices.value" :key="c.key" :value="c.key">{{ c.name }}</option>
        </select>
      </span>
      <span class="xs-tb-group">
        <span v-for="c in COLOURS" :key="c.what" class="xs-tb-split">
          <button type="button" class="xs-tool" :data-tb="`${c.what}-colour`" :title="text(c.title)" :aria-label="text(c.title)" :disabled="!state" @click="paint(c.what, last[c.what])">
            <component :is="c.icon" :colour="last[c.what]" />
          </button>
          <button
            type="button"
            class="xs-tool xs-tb-caret"
            :data-tb="`${c.what}-palette`"
            :title="m.pickColour"
            :aria-label="m.pickColour"
            :aria-expanded="open === c.what"
            :disabled="!state"
            @click="toggle(c.what)"
          >
            <IconDown />
          </button>
          <div v-if="open === c.what" class="xs-tb-pop" role="group" :aria-label="text(c.title)">
            <button type="button" class="xs-tb-item" :data-tb="c.noneHook" :aria-pressed="!shown[c.what]" @click="paint(c.what, null)">{{ text(c.none) }}</button>
            <ColourSwatches :model-value="shown[c.what]" @update:model-value="paint(c.what, $event)" />
          </div>
        </span>
      </span>
      <span class="xs-tb-group">
        <span class="xs-tb-split">
          <button type="button" class="xs-tool xs-tb-wide" data-tb="borders" :title="m.borders" :aria-label="m.borders" :aria-expanded="open === 'borders'" :disabled="!state" @click="toggle('borders')">
            <IconBorderAll />
            <IconDown class="xs-tb-down" />
          </button>
          <div v-if="open === 'borders'" class="xs-tb-pop" role="menu" :aria-label="m.borders">
            <button v-for="b in BORDERS" :key="b.kind" type="button" class="xs-tb-item" role="menuitem" :data-tb="`border-${b.kind}`" @click="border(b.kind)">
              <component :is="b.icon" />
              {{ text(b.title) }}
            </button>
          </div>
        </span>
      </span>
      <span class="xs-tb-group">
        <button
          v-for="a in ALIGN"
          :key="a.h"
          type="button"
          :class="['xs-tool', { 'xs-tool-on': style?.alignment?.horizontal === a.h }]"
          :data-tb="`align-${a.h}`"
          :title="text(a.title)"
          :aria-label="text(a.title)"
          :aria-pressed="style?.alignment?.horizontal === a.h"
          :disabled="!state"
          @click="apply(over({ style: alignToggle(style, a.h) }))"
        >
          <component :is="a.icon" />
        </button>
        <button
          type="button"
          :class="['xs-tool', { 'xs-tool-on': !!style?.alignment?.wrap }]"
          data-tb="wrap"
          :title="m.wrapText"
          :aria-label="m.wrapText"
          :aria-pressed="!!style?.alignment?.wrap"
          :disabled="!state"
          @click="apply(over({ style: wrapToggle(style) }))"
        >
          <IconWrap />
        </button>
      </span>
      <span v-if="sheet && source.structure" class="xs-tb-group">
        <span class="xs-tb-split">
          <button
            type="button"
            :class="['xs-tool', { 'xs-tool-on': merged }]"
            data-tb="merge"
            :title="m.mergeCenter"
            :aria-label="m.mergeCenter"
            :aria-pressed="merged"
            :disabled="!state?.canMerge"
            @click="merge(merged ? 'unmergeCells' : 'mergeCenter')"
          >
            <IconMerge />
          </button>
          <button type="button" class="xs-tool xs-tb-caret" data-tb="merge-menu" :title="m.mergeMenu" :aria-label="m.mergeMenu" :aria-expanded="open === 'merge'" :disabled="!state?.canMerge" @click="toggle('merge')">
            <IconDown />
          </button>
          <div v-if="open === 'merge'" class="xs-tb-pop" role="menu" :aria-label="m.mergeMenu">
            <button v-for="x in MERGES" :key="x" type="button" class="xs-tb-item" role="menuitem" :data-tb="kebab(x)" @click="merge(x)">
              {{ m[x] }}
            </button>
          </div>
        </span>
      </span>
      <span v-if="hostDialog() || sheet" class="xs-tb-group">
        <button type="button" class="xs-tool" data-tb="format-cells" :title="m.formatCellsOpen" :aria-label="m.formatCellsOpen" :disabled="!state" @click="formatCells">
          <IconFormatCells />
        </button>
      </span>
    </div>
  </div>
</template>

<style src="../controls.css"></style>
<style src="../toolbar.css"></style>
