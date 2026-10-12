<script setup lang="ts">
// Excel's formatting toolbar over a sheet source: undo and redo, bold, italic, underline, font and fill colours, borders,
// horizontal alignment, wrapping and the number format of the selection the sheet's `state` event gives, and Format Cells
// for a host that shows the dialog. On a narrow frame what does not fit goes into a menu; on a finger every target is 44 × 44 px.
import { computed, getCurrentInstance, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import type { SheetViewApi } from '../core/commands'
import type { SheetViewState } from '../core/sheet'
import { enMessages, type SheetMessages } from '../messages'
import type { EditOutcome, FormatPatch } from '../source'
import {
  alignToggle,
  borderPatches,
  FIRST_FILL_COLOUR,
  FIRST_FONT_COLOUR,
  fontOn,
  fontToggle,
  numberKindOf,
  PALETTE,
  wrapToggle,
  type BorderKind,
  type FontToggle,
  type HorizontalAlign,
  type NumberKind,
  type ToolbarNumberFormat,
  type ToolbarSource,
} from '../toolbar'
import { IconDown } from './icons'
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
    /** The number formats to offer, with their codes in the book's locale; without them there is no number format list. */
    numberFormats?: readonly ToolbarNumberFormat[]
    /**
     * The sheet (a `SheetView`'s template ref) to format, undo and redo through, as its keys do: it applies the outcome, sets
     * the selection and shows a refusal itself. Without it the toolbar calls the source and emits `select` and `error`.
     */
    sheet?: SheetViewApi | null
  }>(),
  { messages: undefined, numberFormats: undefined, sheet: null },
)
const emit = defineEmits<{
  /** Where the selection goes after an undo or a redo (`EditOutcome.select`), for the sheet's `focus`; not with `sheet`. */
  select: [place: { sheet: string; range: string }]
  /** The source refused or failed a call: its message; not with `sheet`, which shows it. */
  error: [message: string]
  /** A call is done: the host may give the grid its focus back. */
  done: []
  /** Format Cells… pressed: the host opens `FormatDialog`; without a listener there is no such button. */
  formatCells: []
}>()

/** Under this width the toolbar keeps its first buttons in a row and the rest in a menu, as the sheet goes narrow. */
const NARROW_PX = 600

const m = computed<SheetMessages>(() => ({ ...enMessages, ...props.messages }))
const instance = getCurrentInstance()
/** Whether the host listens to `formatCells`, so that the button opens its dialog. */
const hostDialog = typeof instance?.vnode.props?.onFormatCells === 'function' || Array.isArray(instance?.vnode.props?.onFormatCells)
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
const KIND_TITLE: Record<NumberKind, keyof SheetMessages> = {
  general: 'formatGeneral',
  number: 'formatNumber',
  currency: 'formatCurrency',
  percent: 'formatPercent',
  date: 'formatDate',
  time: 'formatTime',
  text: 'formatText',
}
const text = (key: keyof SheetMessages) => m.value[key] as string

// ── what the active cell shows ──
const numberKind = computed(() => numberKindOf(props.state?.numberFormat, props.numberFormats ?? []))
const fontColour = computed(() => style.value?.font?.color?.toUpperCase() ?? null)
const fillColour = computed(() => style.value?.fill?.color?.toUpperCase() ?? null)
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
const open = ref<'font' | 'fill' | 'borders' | null>(null)
const menu = ref(false)
const toggle = (what: 'font' | 'fill' | 'borders') => (open.value = open.value === what ? null : what)
function closeAll(): void {
  open.value = null
  menu.value = false
}
/** A press outside the toolbar closes its menus. */
function onDocDown(e: PointerEvent): void {
  if (root.value && !root.value.contains(e.target as Node)) closeAll()
}
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
  document.addEventListener('pointerdown', onDocDown, true)
  if (typeof window.matchMedia === 'function') {
    coarseQuery = window.matchMedia('(pointer: coarse)')
    onCoarse()
    coarseQuery.addEventListener?.('change', onCoarse)
  }
})
onBeforeUnmount(() => {
  unsubscribe?.()
  observer?.disconnect()
  document.removeEventListener('pointerdown', onDocDown, true)
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
const lastFont = ref(FIRST_FONT_COLOUR)
const lastFill = ref(FIRST_FILL_COLOUR)
function paintFont(colour: string | null): void {
  if (colour) lastFont.value = colour
  open.value = null
  apply(over({ style: { font: { color: colour } } }))
}
function paintFill(colour: string | null): void {
  if (colour) lastFill.value = colour
  open.value = null
  apply(over({ style: { fill: colour ? { color: colour } : null } }))
}
function border(kind: BorderKind): void {
  open.value = null
  if (props.state) apply(borderPatches(props.state.ref, kind))
}
function pickNumber(e: Event): void {
  const f = props.numberFormats?.find((x) => x.kind === (e.target as HTMLSelectElement).value)
  if (f) apply(over({ numberFormat: f.format }))
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
      <span v-if="numberFormats?.length" class="xs-tb-group">
        <select class="xs-tb-select" data-tb="number" :title="m.numberFormat" :aria-label="m.numberFormat" :value="numberKind ?? ''" :disabled="!state" @change="pickNumber">
          <option v-if="!numberKind" value="" disabled>{{ m.formatOther }}</option>
          <option v-for="f in numberFormats" :key="f.kind" :value="f.kind">{{ text(KIND_TITLE[f.kind]) }}</option>
        </select>
      </span>
      <span class="xs-tb-group">
        <span class="xs-tb-split">
          <button type="button" class="xs-tool" data-tb="font-colour" :title="m.fontColour" :aria-label="m.fontColour" :disabled="!state" @click="paintFont(lastFont)">
            <IconFontColour :colour="lastFont" />
          </button>
          <button type="button" class="xs-tool xs-tb-caret" data-tb="font-palette" :title="m.pickColour" :aria-label="m.pickColour" :aria-expanded="open === 'font'" :disabled="!state" @click="toggle('font')">
            <IconDown />
          </button>
          <div v-if="open === 'font'" class="xs-tb-pop" role="group" :aria-label="m.fontColour">
            <button type="button" class="xs-tb-item" data-tb="colour-auto" :aria-pressed="!fontColour" @click="paintFont(null)">{{ m.colourAuto }}</button>
            <div class="xs-tb-swatches">
              <button
                v-for="(c, i) in PALETTE"
                :key="i"
                type="button"
                class="xs-tb-swatch"
                :style="{ backgroundColor: c }"
                :title="c"
                :aria-label="c"
                :data-tb-colour="c"
                :aria-pressed="fontColour === c"
                @click="paintFont(c)"
              />
            </div>
          </div>
        </span>
        <span class="xs-tb-split">
          <button type="button" class="xs-tool" data-tb="fill-colour" :title="m.fillColour" :aria-label="m.fillColour" :disabled="!state" @click="paintFill(lastFill)">
            <IconFillColour :colour="lastFill" />
          </button>
          <button type="button" class="xs-tool xs-tb-caret" data-tb="fill-palette" :title="m.pickColour" :aria-label="m.pickColour" :aria-expanded="open === 'fill'" :disabled="!state" @click="toggle('fill')">
            <IconDown />
          </button>
          <div v-if="open === 'fill'" class="xs-tb-pop" role="group" :aria-label="m.fillColour">
            <button type="button" class="xs-tb-item" data-tb="fill-none" :aria-pressed="!fillColour" @click="paintFill(null)">{{ m.fillNone }}</button>
            <div class="xs-tb-swatches">
              <button
                v-for="(c, i) in PALETTE"
                :key="i"
                type="button"
                class="xs-tb-swatch"
                :style="{ backgroundColor: c }"
                :title="c"
                :aria-label="c"
                :data-tb-colour="c"
                :aria-pressed="fillColour === c"
                @click="paintFill(c)"
              />
            </div>
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
      <span v-if="hostDialog" class="xs-tb-group">
        <button type="button" class="xs-tool" data-tb="format-cells" :title="m.formatCellsOpen" :aria-label="m.formatCellsOpen" :disabled="!state" @click="emit('formatCells')">
          <IconFormatCells />
        </button>
      </span>
    </div>
  </div>
</template>

<style src="../toolbar.css"></style>
