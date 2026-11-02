<script setup lang="ts">
// Excel's formatting toolbar over a sheet source, in Excel for Windows' ribbon order: undo and redo, Format Painter (through the
// sheet), the font and its size, bold, italic, Underline with its menu, borders, fill and font colors, vertical alignment, Orientation
// with its menu, horizontal alignment, indent, wrapping, Merge & Center with its dropdown (through the sheet, with the source's `structure`), Excel's Number Format box (xlcalc's,
// in the workbook's locale) and its buttons, Accounting Number Format with its currencies, applied to the selection the sheet's `state` event reports, Format Cells for a host
// that displays the dialog, and AutoSum with its list (through the sheet, with the source's `autoSum`); then Insert > PivotTable
// (its dialog over the sheet), and Data > Refresh All with its list's Refresh of the PivotTable under the active cell (through the
// sheet, with the source's `refreshConnections`); `items` picks and orders the groups and adds the host's buttons. In a narrow
// viewport what does not fit goes into a menu; on a touch screen every target is 44 × 44 px.
import { computed, getCurrentInstance, h, nextTick, onBeforeUnmount, onBeforeUpdate, onMounted, reactive, ref, watch, type FunctionalComponent } from 'vue'
import type { FormatCellsAt, MergeCommand, SheetViewApi, SheetViewState, ToolbarButton, ToolbarGroup, ToolbarItem, ToolbarSource } from '../contract/view'
import { failedOutcome } from '../failure'
import { NARROW_PX, useCoarsePointer } from '../frame'
import { useMessages } from '../messages'
import type { SheetMessages } from '../contract/messages'
import { useNumberFormats } from '../number-formats'
import { usableTag } from '../locales'
import { allows } from '../protection'
import type { AutoSumFunction, EditOutcome, FormatPatch } from '../contract/source'
import {
  alignToggle,
  borderPatches,
  FIRST_FILL_COLOUR,
  FIRST_FONT_COLOUR,
  FONT_NAMES,
  FONT_SIZES,
  fontOn,
  fontSizeOf,
  fontSizeStep,
  fontToggle,
  indentStep,
  ORIENTATIONS,
  orientationOn,
  orientationToggle,
  underlineOn,
  underlineToggle,
  ACCOUNTING_STYLE,
  COMMA_STYLE,
  decimalStep,
  fitLines,
  fitRow,
  PERCENT_STYLE,
  ribbonGroups,
  shownFont,
  TOOLBAR_GROUPS,
  upperColour,
  verticalOn,
  verticalToggle,
  wrapToggle,
  type BookFont,
  type BorderKind,
  type FontToggle,
  type HorizontalAlign,
  type OrientationKind,
  type RibbonGroup,
  type RibbonPiece,
  type RibbonSlot,
  type UnderlineKind,
  type VerticalAlign,
} from '../toolbar'
import type { StylePatch } from 'xlcalc/contract'
import ColourSwatches from './ColourSwatches.vue'
import ToolbarBox from './ToolbarBox.vue'
import ToolbarPivot from './ToolbarPivot.vue'
import { ToolbarLabel } from './ToolbarLabel'
import ToolbarDrop, { type ToolbarDropItem } from './ToolbarDrop.vue'
import { hasListener } from './dom'
import { IconDown } from './icons'
import { useOutsidePress } from './popup'
import { IconPivotTable } from './pivot-icons'
import {
  IconAlignBottom,
  IconAlignCenter,
  IconAlignLeft,
  IconAlignMiddle,
  IconAlignRight,
  IconAlignTop,
  IconAngleDown,
  IconAngleUp,
  IconAutoSum,
  IconRefreshAll,
  IconBold,
  IconBorderAll,
  IconBorderBottom,
  IconBorderNone,
  IconBorderOutside,
  IconFillColour,
  IconFontColour,
  IconFontDown,
  IconFontUp,
  IconFormatCells,
  IconFormatPainter,
  IconIndentDecrease,
  IconIndentIncrease,
  IconItalic,
  IconMerge,
  IconMore,
  IconOrientation,
  IconRotateDown,
  IconRotateUp,
  IconTextVertical,
  IconAccounting,
  IconComma,
  IconDecimalDown,
  IconDecimalUp,
  IconPercent,
  IconRedo,
  IconUnderline,
  IconUnderlineDouble,
  IconUndo,
  IconWrap,
} from './toolbar-icons'
import '../controls.css'

const props = withDefaults(
  defineProps<{
    source: ToolbarSource
    /** The payload of the sheet's `state` event: the selection and the active cell's formatting; with null the buttons are disabled. */
    state: SheetViewState | null
    /** Interface texts; a partial set overrides the English defaults. */
    messages?: Partial<SheetMessages>
    /**
     * The workbook's locale, in which the Number Format box names its categories and items, in place of the sheet's
     * (`state.locale`) or else the source's `locale`, which it is by default; en-US without any.
     */
    locale?: string
    /**
     * The sheet (a `SheetView`'s template ref) to format, undo and redo through, as its keyboard commands do: it applies the
     * outcome, sets the selection and displays a rejection itself. Without it the toolbar calls the source and emits `select`
     * and `error`.
     */
    sheet?: SheetViewApi | null
    /** The host opens its own dialog on `formatCells`: for a host whose listener the toolbar cannot see, such as a custom element's. */
    hostDialog?: boolean
    /** The Font box's list, in place of Excel's common fonts: a browser does not tell which fonts are installed. */
    fonts?: readonly string[]
    /** The groups to draw, in order, and the host's buttons among them; all groups in their default order (`TOOLBAR_GROUPS`) without it. */
    items?: readonly ToolbarItem[]
    /** How many lines the items may take at 600 px and wider before the rest go into the menu under More; one below 600 px. 2 is Excel's ribbon: its groups two rows tall. */
    rows?: number
  }>(),
  { messages: undefined, locale: undefined, sheet: null, hostDialog: false, fonts: () => FONT_NAMES, items: () => TOOLBAR_GROUPS, rows: 1 },
)
const emit = defineEmits<{
  /** Where the selection goes after an undo or a redo (`EditOutcome.select`), for the sheet's `focus`; not emitted with `sheet`. */
  select: [place: { sheet: string; range: string }]
  /** The source rejected a call or the call failed: its message; not emitted with `sheet`, which displays it itself. */
  error: [message: string]
  /** A call has completed, or Esc left the Font or Font Size box: the host may return focus to the grid. */
  done: []
  /**
   * Format Cells… was clicked, or an item that opens it (More Accounting Formats..., Format Cell Alignment) with where it opens, for a
   * host with its own `FormatDialog`; without a listener these open the sheet's dialog (through `sheet`), and without either they are absent.
   */
  formatCells: [at?: FormatCellsAt]
}>()

const m = useMessages(() => props.messages)
const instance = getCurrentInstance()
/** Whether the host currently opens its own dialog on `formatCells`, instead of the sheet's. */
const dialogToHost = () => props.hostDialog || hasListener(instance, 'onFormatCells')
/** Whether the host listens, read again before each render: a listener is not reactive, as a template condition it was read on each render. */
const listening = ref(dialogToHost())
onBeforeUpdate(() => (listening.value = dialogToHost()))
function formatCells(at?: FormatCellsAt): void {
  closeAll()
  if (dialogToHost() || !props.sheet) emit('formatCells', at)
  else props.sheet.formatCells(at)
}
/** Whether Format Cells can open: the sheet's dialog, or the host's. */
const dialogShown = computed(() => !!props.sheet || listening.value)
const style = computed(() => props.state?.style)
/** The formatting buttons work on a sheet whose protection, if any, allows formatting cells, as Excel's are disabled otherwise. */
const formattable = computed(() => !!props.state && allows(props.state.protection, 'formatCells'))

const FONT: { what: FontToggle; icon: typeof IconBold; title: keyof SheetMessages }[] = [
  { what: 'bold', icon: IconBold, title: 'bold' },
  { what: 'italic', icon: IconItalic, title: 'italic' },
]
const VERTICAL: { v: VerticalAlign; icon: typeof IconBold; title: keyof SheetMessages }[] = [
  { v: 'top', icon: IconAlignTop, title: 'alignTop' },
  { v: 'center', icon: IconAlignMiddle, title: 'alignMiddle' },
  { v: 'bottom', icon: IconAlignBottom, title: 'alignBottom' },
]
const ALIGN: { h: HorizontalAlign; icon: typeof IconBold; title: keyof SheetMessages }[] = [
  { h: 'left', icon: IconAlignLeft, title: 'alignLeft' },
  { h: 'center', icon: IconAlignCenter, title: 'alignCenter' },
  { h: 'right', icon: IconAlignRight, title: 'alignRight' },
]
/** Excel's Borders menu, the items it has here in its order: Bottom Border, then after its line No Border, All Borders, Outside Borders. */
const BORDERS: { kind: BorderKind; icon: typeof IconBold; title: keyof SheetMessages }[] = [
  { kind: 'bottom', icon: IconBorderBottom, title: 'borderBottom' },
  { kind: 'none', icon: IconBorderNone, title: 'borderNone' },
  { kind: 'all', icon: IconBorderAll, title: 'borderAll' },
  { kind: 'outside', icon: IconBorderOutside, title: 'borderOutside' },
]
/** Excel's dropdown under Merge & Center: each item is named by its command's text, and its hook (`data-tb`) is the command in kebab case. */
const MERGES: readonly MergeCommand[] = ['mergeCenter', 'mergeAcross', 'mergeCells', 'unmergeCells']
/** Excel's AutoSum list in its order; the source's `autoSumFunctions` says which besides Sum it builds. */
const AUTO_SUMS: { fn: AutoSumFunction; title: keyof SheetMessages }[] = [
  { fn: 'sum', title: 'autoSumSum' },
  { fn: 'average', title: 'autoSumAverage' },
  { fn: 'count', title: 'autoSumCount' },
  { fn: 'max', title: 'autoSumMax' },
  { fn: 'min', title: 'autoSumMin' },
]
const kebab = (name: string) => name.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)
/** Excel's split buttons for fill and font color, in its Font group's order: the button applies the last color, the menu offers the no-color choice and the palette. */
type Colouring = 'font' | 'fill'
const COLOURS: { what: Colouring; icon: typeof IconFontColour; title: keyof SheetMessages; none: keyof SheetMessages; noneHook: string }[] = [
  { what: 'fill', icon: IconFillColour, title: 'fillColor', none: 'fillNone', noneHook: 'fill-none' },
  { what: 'font', icon: IconFontColour, title: 'fontColor', none: 'colorAuto', noneHook: 'colour-auto' },
]
const text = (key: keyof SheetMessages) => m.value[key] as string

// ── the active cell's formatting ──
/** Excel's box: the list's items, and the name of the category of the active cell's format (General without one). */
/** The workbook's locale: the prop, else the sheet's, else the source's. */
const locale = computed(() => props.locale ?? props.state?.locale ?? props.source.locale)
const numbers = useNumberFormats(() => locale.value)
const category = computed(() => (props.state ? numbers.categoryOf(props.state.numberFormat) : null))
/** The workbook's default font, for a cell without its own: the source does not report it yet, so it plugs in here when it does. */
/** The workbook's default font, which a cell without its own font name or size displays. */
const book = (): BookFont => props.state?.font ?? {}
const font = computed(() => shownFont(style.value, book()))
const sizes = FONT_SIZES.map(String)
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

// ── the viewport: narrow, touch ──
const root = ref<HTMLElement | null>(null)
const width = ref(0)
const coarse = useCoarsePointer()
/** Below `NARROW_PX` the toolbar's targets are touch-sized; at any width its row keeps the items that fit and the rest go into a menu. */
const narrow = computed(() => (width.value || (typeof window === 'undefined' ? NARROW_PX : window.innerWidth)) < NARROW_PX)
const touch = computed(() => narrow.value || coarse.value)
const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(([entry]) => (width.value = entry?.contentRect.width ?? 0))
watch(root, (el, old) => {
  if (old) observer?.unobserve(old)
  if (el) observer?.observe(el)
})

// ── the items: groups in the host's order, its buttons together between them ──
/** A group, or a run of the host's buttons with the same `group`. */
type Block = { key: string; group: ToolbarGroup; buttons?: undefined } | { key: string; group?: undefined; buttons: ToolbarButton[] }
/** The active cell is in a PivotTable: a boolean, so that the blocks change only when it does. */
const inPivot = computed(() => !!props.state?.pivot)
/** Whether a group is drawn: some need the sheet or the source's methods. */
function drawn(group: ToolbarGroup): boolean {
  const src = props.source
  const sheet = !!props.sheet
  switch (group) {
    case 'undo':
      return !!(src.undo || src.redo)
    case 'formatPainter':
    case 'autoSum':
      return sheet
    case 'merge':
    case 'pivotTable':
      return sheet && !!src.structure
    case 'formatCells':
      return dialogShown.value
    case 'refreshAll':
      return sheet && !!src.refreshConnections
    case 'pivotAnalyze':
    case 'pivotDesign':
      return sheet && inPivot.value && !!src.pivotCells
    default:
      return true
  }
}
const blocks = computed(() => {
  const out: Block[] = []
  const seen = new Set<string>()
  for (const item of props.items) {
    const last = out[out.length - 1]
    if (typeof item !== 'string') {
      if (last?.buttons && last.buttons[0]!.group === item.group) last.buttons.push(item)
      else out.push({ key: `host-${out.length}`, buttons: [item] })
    } else if (TOOLBAR_GROUPS.includes(item) && !seen.has(item)) {
      seen.add(item)
      if (drawn(item)) out.push({ key: item, group: item })
    }
  }
  return out
})
/** The lines the row may take: `rows` (at least one, whole) at 600 px and wider, one below. */
const lines = computed(() => (narrow.value ? 1 : Math.max(1, Math.floor(props.rows) || 1)))
/** Two lines are Excel's ribbon: one sequence of groups, each two rows tall, its controls in Excel's rows. */
const ribbon = computed(() => lines.value === 2)
/** The layout of the groups: Excel's two-row groups, Excel's groups on one line, or on a phone each block alone. */
const layout = computed(() => (ribbon.value ? 'ribbon' : narrow.value ? 'blocks' : 'groups'))
/** The groups drawn between separators, fitted into the row whole. */
const groups = computed(() => ribbonGroups(blocks.value, layout.value))
/** The keys of the groups that stay in the row, as measured; null while measuring, when all are drawn there. */
const kept = ref<ReadonlySet<string> | null>(null)
/** The keys of the kept groups that start a line after the first, drawn without the separator before them. */
const breaks = ref<ReadonlySet<string>>(new Set())
/** The row on its lines and the menu with the blocks of the groups that do not fit on them, each on one line (no menu when all fit). */
const parts = computed(() => {
  const all = groups.value
  const on = kept.value
  if (!on || all.every((g) => on.has(g.key))) return [{ place: 'row', groups: all }]
  // The menu keeps the order of `items`, whatever places the ribbon gave the blocks
  const out = new Set(all.filter((g) => !on.has(g.key)).flatMap((g) => g.stacks.flatMap((st) => st.lines.flat().map((p) => p.block))))
  const rest = blocks.value.filter((b) => out.has(b))
  return [
    { place: 'row', groups: all.filter((g) => on.has(g.key)) },
    { place: 'menu', groups: ribbonGroups(rest, narrow.value ? 'blocks' : 'groups') },
  ]
})
/** The More button: while the row is measured too, so that its width counts. */
const moreShown = computed(() => parts.value.length > 1 || kept.value === null)
/**
 * Measures the row drawn with every group. On one line (and in the ribbon's one sequence) the groups that fit beside More stay
 * in order, a group too wide for what is left going into More while later ones still take the row; on more lines all stay when
 * they fit, otherwise those from the start that fit with More at the end of the last line.
 */
function measureRow(): void {
  const row = root.value?.querySelector<HTMLElement>('.xs-tb-row')
  if (!row || kept.value !== null) return
  const widthOf = (el: Element | null) => el?.getBoundingClientRect().width ?? 0
  const keys = groups.value.map((g) => g.key)
  const widths = keys.map((key) => widthOf(row.querySelector(`:scope > [data-tb-block="${key}"]`)))
  const more = widthOf(row.querySelector('[data-tb="more"]'))
  if (ribbon.value || lines.value === 1) {
    breaks.value = new Set()
    kept.value = new Set(fitRow(widths, row.clientWidth, more).map((i) => keys[i]!))
    return
  }
  const out = fitLines(widths, row.clientWidth, lines.value, more)
  breaks.value = new Set(out.breaks.map((i) => keys[i]!))
  kept.value = new Set(keys.slice(0, out.fit))
}
/** Whether a piece of a block draws the controls of this slot: all of them where the block is drawn whole. */
const has = (piece: RibbonPiece<Block>, slot: RibbonSlot) => !piece.slot || piece.slot === slot
// Whatever changes the row's widths or lines draws it whole again and measures it: the sheet and the source decide which
// groups are drawn, and the Number Format box widens when its locale's list loads
watch(
  [narrow, width, touch, groups, lines, () => props.messages, () => props.sheet, () => props.source, () => props.hostDialog, () => numbers.choices.value.length > 0],
  () => {
    kept.value = null
    breaks.value = new Set()
    void nextTick(measureRow)
  },
  { flush: 'post' },
)
/** The hook of a block that starts a line after the first. */
const lineStart = (key: string) => (breaks.value.has(key) ? '' : undefined)
/** A host's button: its action, after the menus close. */
function runHost(button: { run?: () => void }): void {
  closeAll()
  button.run?.()
}
/** A host's button drawn large: `size: 'large'` in a group stacked two to a column (in the ribbon). */
const large = (g: RibbonGroup<Block>, x: ToolbarButton) => g.kind === 'cols' && x.size === 'large'
/** No wrapper: a host's button without a list is drawn as its button alone. */
const Bare: FunctionalComponent = (_, { slots }) => slots.default?.()
/** A host's button's icon as the toolbar draws its own: a component, SVG markup, or without an icon its label as text. */
const HostIcon: FunctionalComponent<{ button: ToolbarButton }> = ({ button: b }) =>
  typeof b.icon === 'string' ? h('span', { class: 'xs-icon xs-tb-svg', innerHTML: b.icon }) : b.icon ? h(b.icon, { class: 'xs-icon' }) : b.label
HostIcon.props = ['button']

// ── ScreenTips: a button's name (and key) below it, sooner than the browser's title; mouse only ──
/** How long the pointer rests on a button before its tip shows, and how long after a tip hides the next one shows at once, as Office's do. */
const TIP_DELAY = 400
const TIP_WARM = 600
const tip = ref<{ el: HTMLElement; text: string; more?: string; left: number; top: number } | null>(null)
const tipEl = ref<HTMLElement | null>(null)
let tipTimer: ReturnType<typeof setTimeout> | undefined
let warmUntil = 0
function hideTip(warm = true): void {
  clearTimeout(tipTimer)
  if (tip.value && warm) warmUntil = Date.now() + TIP_WARM
  tip.value = null
}
function showTip(el: HTMLElement): void {
  const text = el.dataset.tip
  if (!text || open.value !== null || !el.isConnected) return
  const r = el.getBoundingClientRect()
  tip.value = { el, text, more: el.dataset.tipMore, left: r.left, top: r.bottom + 4 }
  // Kept inside the viewport: a tip near the right edge moves left
  void nextTick(() => {
    const box = tipEl.value?.getBoundingClientRect()
    if (box && tip.value && box.right > window.innerWidth - 4) tip.value = { ...tip.value, left: Math.max(4, window.innerWidth - 4 - box.width) }
  })
}
function onTipOver(e: PointerEvent): void {
  if (e.pointerType !== 'mouse') return
  const el = (e.target as HTMLElement).closest<HTMLElement>('[data-tip]')
  if (!el || el === tip.value?.el) return
  hideTip()
  if (Date.now() < warmUntil) showTip(el)
  else tipTimer = setTimeout(() => showTip(el), TIP_DELAY)
}
function onTipOut(e: PointerEvent): void {
  const from = (e.target as HTMLElement).closest('[data-tip]')
  if (from && !from.contains(e.relatedTarget as Node | null)) hideTip()
}

// ── menus: the color palettes, the borders, and the narrow layout's menu ──
/** The toolbar's lists, and a host's button's (`host:<id>`). */
type Menu = 'font' | 'fill' | 'borders' | 'underline' | 'accounting' | 'orientation' | 'merge' | 'autoSum' | 'refreshAll' | 'fontName' | 'fontSize' | `host:${string}`
const open = ref<Menu | null>(null)
const menu = ref(false)
const toggle = (what: Menu) => {
  hideTip(false)
  open.value = open.value === what ? null : what
}
/** A box's list opens or closes, as the toolbar's other menus do. */
const listOpen = (what: Menu, on: boolean) => (open.value = on ? what : open.value === what ? null : open.value)
function closeAll(): void {
  open.value = null
  menu.value = false
}
useOutsidePress(root, closeAll, () => open.value !== null || menu.value)
/**
 * Esc closes an open menu wherever the focus is, as Excel's does: the toolbar's buttons leave it on the grid, which then does
 * not see this Esc. An Esc within a toolbar is that toolbar's own (a box restores what it showed).
 */
function escOutside(e: KeyboardEvent): void {
  if (e.key !== 'Escape' || !root.value?.isConnected || (e.target as Element | null)?.closest?.('.xs-toolbar')) return
  e.preventDefault()
  e.stopPropagation()
  closeAll()
}
watch(
  () => open.value !== null || menu.value,
  (on) => (on ? document.addEventListener('keydown', escOutside, true) : document.removeEventListener('keydown', escOutside, true)),
)
/** Buttons do not take focus from the grid, so its keys keep working; the number format list and the font boxes do take it. */
function keepFocus(e: MouseEvent): void {
  if (!(e.target as HTMLElement).closest('select, input')) e.preventDefault()
}

let unsubscribe: (() => void) | null = null
function subscribe(): void {
  unsubscribe?.()
  unsubscribe = props.source.format ? (props.source.subscribe?.(() => tick.value++) ?? null) : null
}
watch(() => props.source, subscribe)
onMounted(() => {
  subscribe()
  measureRow()
})
onBeforeUnmount(() => {
  document.removeEventListener('keydown', escOutside, true)
  clearTimeout(tipTimer)
  unsubscribe?.()
  observer?.disconnect()
})

// ── calls: through the sheet when there is one, otherwise the source's own ──
async function run(call: () => Promise<EditOutcome | null>): Promise<EditOutcome | null> {
  let out: EditOutcome | null
  try {
    out = await call()
  } catch (e) {
    out = failedOutcome(e)
  }
  tick.value++
  if (out && !out.ok && !props.sheet) emit('error', out.message || m.value.editFailed)
  emit('done')
  return out
}
/** Patches over the selection's sheet, in one call and so a single undo step. */
function apply(patches: FormatPatch[]): void {
  const st = props.state
  const format = props.source.format
  const sheet = props.sheet
  if (!st || !format || !patches.length) return
  void run(() => (sheet ? sheet.format(patches) : format.call(props.source, st.sheet, patches)))
}
/** The selection's ranges: every one of several, as Excel formats them all. */
const ranges = (): string[] => (props.state ? (props.state.areas ?? [props.state.ref]) : [])
const over = (p: Omit<FormatPatch, 'range'>): FormatPatch[] => ranges().map((range) => ({ range, ...p }))
/** A style patch over the selection; nothing for none (a step at its limit). */
const restyle = (p: StylePatch | null) => p && apply(over({ style: p }))
/** Increase or Decrease Decimal: the active cell's next format over the selection; nothing when the step changes nothing. */
function stepDecimal(up: boolean): void {
  const next = props.state && decimalStep(props.state, up, usableTag(locale.value))
  if (next !== null && next !== undefined) apply(over({ numberFormat: next }))
}

// The Font and Font Size boxes: a name as typed or picked; a size of 1 to 409, otherwise the box keeps it marked invalid.
const checkSize = (text: string) => (fontSizeOf(text) === null ? m.value.fmtSizeRange : null)
function setFont(name: string): void {
  open.value = null
  restyle({ font: { name } })
}
function setSize(text: string): void {
  open.value = null
  const size = fontSizeOf(text)
  if (size !== null) restyle({ font: { size } })
}

/** Undo and redo one at a time: a second press while one is pending sends nothing. */
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

// The split buttons apply the color picked last, as Excel's do.
const last = reactive<Record<Colouring, string>>({ font: FIRST_FONT_COLOUR, fill: FIRST_FILL_COLOUR })
const shown = computed<Record<Colouring, string | null>>(() => ({ font: fontColour.value, fill: fillColour.value }))
/** A color, or none (Automatic, No Fill), on the text or the fill of the selection. */
function paint(what: Colouring, colour: string | null): void {
  if (colour) last[what] = colour
  open.value = null
  apply(over({ style: what === 'font' ? { font: { color: colour } } : { fill: colour ? { color: colour } : null } }))
}
/** The Borders split button's part applies the border picked last, Bottom Border at first, and is named after it, as Excel's. */
const lastBorder = ref<(typeof BORDERS)[number]>(BORDERS[0]!)
function border(kind: BorderKind): void {
  open.value = null
  lastBorder.value = BORDERS.find((b) => b.kind === kind) ?? lastBorder.value
  apply(ranges().flatMap((range) => borderPatches(range, kind)))
}
/** Underline's split button applies the underline picked last from its menu, Underline at first, and is named after it, as Excel's. */
const lastUnderline = ref<UnderlineKind>('single')
const UNDERLINES: { kind: UnderlineKind; icon: typeof IconBold; title: keyof SheetMessages }[] = [
  { kind: 'single', icon: IconUnderline, title: 'underlineSingle' },
  { kind: 'double', icon: IconUnderlineDouble, title: 'underlineDouble' },
]
const underlinePart = computed(() => (lastUnderline.value === 'double' ? { icon: IconUnderlineDouble, title: m.value.underlineDouble } : { icon: IconUnderline, title: m.value.underline }))
function underline(kind: UnderlineKind): void {
  open.value = null
  lastUnderline.value = kind
  restyle(underlineToggle(style.value, kind))
}
const underlineItems = computed<ToolbarDropItem[]>(() =>
  UNDERLINES.map((u) => ({ hook: `underline-${u.kind}`, label: text(u.title), icon: u.icon, checked: underlineOn(style.value, u.kind), run: () => underline(u.kind) })),
)
/** Accounting Number Format's menu: xlcalc's currencies for the workbook's locale, named in its language, then More Accounting Formats... */
const accountingItems = computed<ToolbarDropItem[]>(() => [
  ...numbers.currencies.value.map(({ currency, name, code }) => ({
    hook: `accounting-${currency.toLowerCase()}`,
    label: name,
    run: () => {
      open.value = null
      apply(over({ numberFormat: code }))
    },
  })),
  ...(dialogShown.value ? [{ hook: 'accounting-more', label: m.value.accountingMore, line: true, run: () => formatCells({ tab: 'number', choice: 'accounting' }) }] : []),
])
/** Excel's Orientation menu: each item checked when the active cell's text is turned so, then Format Cell Alignment. */
const ORIENT_ICONS: Record<OrientationKind, typeof IconBold> = { counterclockwise: IconAngleUp, clockwise: IconAngleDown, vertical: IconTextVertical, up: IconRotateUp, down: IconRotateDown }
const orientationItems = computed<ToolbarDropItem[]>(() => [
  ...ORIENTATIONS.map(({ kind, rotation }) => ({
    hook: `orientation-${kind}`,
    label: m.value.orientationKind(kind),
    icon: ORIENT_ICONS[kind],
    checked: orientationOn(style.value, rotation),
    run: () => {
      open.value = null
      restyle(orientationToggle(style.value, rotation))
    },
  })),
  ...(dialogShown.value ? [{ hook: 'orientation-format', label: m.value.orientationFormat, icon: IconFormatCells, line: true, run: () => formatCells({ tab: 'alignment' }) }] : []),
])
/** Excel's Merge & Center shows as pressed on a merged cell, and clicking it then unmerges. */
const merged = computed(() => !!props.state?.merged)
/** A merge command through the sheet, which warns before values are lost and displays a rejection. */
function merge(command: MergeCommand): void {
  open.value = null
  const sheet = props.sheet
  if (sheet && props.state?.canMerge) void run(() => sheet.merge(command))
}
/** The AutoSum list: Sum and the functions the source builds besides it. */
const autoSums = computed(() => AUTO_SUMS.filter((x) => x.fn === 'sum' || props.source.autoSumFunctions?.includes(x.fn)))
/**
 * Excel's Format Painter through the sheet: a click takes the selection's formats for one painting, or puts a held brush down;
 * a double click (after its two clicks) keeps painting until Esc or the button again.
 */
function paintFormats(keep: boolean): void {
  closeAll()
  const sheet = props.sheet
  if (!sheet) return
  sheet.formatPainter(keep ? 'keep' : props.state?.painting ? 'off' : 'once')
}
/** AutoSum through the sheet, which opens the active cell's edit with the formula; the menus close, as the edit takes focus. */
function autoSum(fn: AutoSumFunction): void {
  closeAll()
  if (props.state?.canAutoSum) void props.sheet?.autoSum(fn)
}
/** Excel's Data > Refresh All through the sheet: every call outside the workbook asked again, not a step of undo. */
function refreshAll(): void {
  closeAll()
  const sheet = props.sheet
  if (sheet && props.state?.canRefreshAll) void run(() => sheet.refreshAll())
}
/** The list's Refresh: Excel's Refresh (Alt+F5) of the PivotTable under the active cell through the sheet, one step of undo. */
function refreshPivot(): void {
  closeAll()
  const sheet = props.sheet
  if (sheet && props.state?.canRefreshPivot) void run(() => sheet.refreshPivot())
}
/** Excel's Insert > PivotTable: its dialog over the sheet. */
function insertPivot(): void {
  closeAll()
  if (props.state?.canInsertPivot) void props.sheet?.insertPivotTable()
}
/** More Functions…: Excel's Insert Function dialog over the sheet. */
function moreFunctions(): void {
  closeAll()
  props.sheet?.insertFunction()
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
    @pointerover="onTipOver"
    @pointerout="onTipOut"
    @pointerdown="hideTip(false)"
  >
    <div
      v-for="part in parts"
      v-show="part.place === 'row' || menu"
      :key="part.place"
      :class="part.place === 'row' ? ['xs-tb-row', { 'xs-tb-rows': lines > 1 && !ribbon, 'xs-tb-ribbon': ribbon }] : 'xs-tb-menu'"
      :data-tb-menu="part.place === 'menu' ? '' : undefined"
    >
      <span v-for="g in part.groups" :key="g.key" :class="['xs-tb-group', `xs-tb-g-${g.kind}`]" :data-tb-block="g.key" :data-tb-line="lineStart(g.key)">
        <span v-for="st in g.stacks" :key="st.key" class="xs-tb-stack">
          <span v-for="(ln, li) in st.lines" :key="li" class="xs-tb-line">
            <template v-for="p in ln" :key="p.block.key">
              <span v-if="p.block.group === 'undo' && (source.undo || source.redo)" class="xs-tb-part">
                <button v-if="source.undo" type="button" class="xs-tool" data-tb="undo" :data-tip="m.undo" :aria-label="m.undo" :disabled="!canUndo" @click="step('undo')">
                  <IconUndo />
                </button>
                <button v-if="source.redo" type="button" class="xs-tool" data-tb="redo" :data-tip="m.redo" :aria-label="m.redo" :disabled="!canRedo" @click="step('redo')">
                  <IconRedo />
                </button>
              </span>
              <span v-else-if="p.block.group === 'formatPainter' && sheet" class="xs-tb-part">
                <button
                  type="button"
                  :class="['xs-tool', { 'xs-tool-on': !!state?.painting }]"
                  data-tb="format-painter"
                  :data-tip="m.formatPainter"
                  :aria-label="m.formatPainter"
                  :aria-pressed="!!state?.painting"
                  :disabled="!state?.canFormatPainter"
                  @click="paintFormats(false)"
                  @dblclick="paintFormats(true)"
                >
                  <IconFormatPainter />
                </button>
              </span>
              <span v-else-if="p.block.group === 'font'" class="xs-tb-part">
                <span class="xs-tb-pair">
                  <ToolbarBox
                    :value="font.name"
                    :items="fonts"
                    :label="m.fontName"
                    hook="font-name"
                    field-class="xs-tb-font"
                    preview
                    complete
                    :open="open === 'fontName'"
                    :disabled="!formattable"
                    @update:open="listOpen('fontName', $event)"
                    @apply="setFont"
                    @leave="emit('done')"
                  />
                  <ToolbarBox
                    :value="font.size === undefined ? '' : String(font.size)"
                    :items="sizes"
                    :label="m.fontSize"
                    hook="font-size"
                    field-class="xs-tb-size"
                    inputmode="decimal"
                    :check="checkSize"
                    :open="open === 'fontSize'"
                    :disabled="!formattable"
                    @update:open="listOpen('fontSize', $event)"
                    @apply="setSize"
                    @leave="emit('done')"
                  />
                </span>
                <span class="xs-tb-pair">
                  <button type="button" class="xs-tool" data-tb="font-up" :data-tip="m.fontSizeUp" :aria-label="m.fontSizeUp" :disabled="!formattable" @click="restyle(fontSizeStep(style, true, book()))">
                    <IconFontUp />
                  </button>
                  <button type="button" class="xs-tool" data-tb="font-down" :data-tip="m.fontSizeDown" :aria-label="m.fontSizeDown" :disabled="!formattable" @click="restyle(fontSizeStep(style, false, book()))">
                    <IconFontDown />
                  </button>
                </span>
              </span>
              <span v-else-if="p.block.group === 'fontStyle'" class="xs-tb-part">
                <button
                  v-for="b in FONT"
                  :key="b.what"
                  type="button"
                  :class="['xs-tool', { 'xs-tool-on': fontOn(style, b.what) }]"
                  :data-tb="b.what"
                  :data-tip="text(b.title)"
                  :aria-label="text(b.title)"
                  :aria-pressed="fontOn(style, b.what)"
                  :disabled="!formattable"
                  @click="apply(over({ style: fontToggle(style, b.what) }))"
                >
                  <component :is="b.icon" />
                </button>
                <span class="xs-tb-split">
                  <button
                    type="button"
                    :class="['xs-tool', { 'xs-tool-on': underlineOn(style, lastUnderline) }]"
                    data-tb="underline"
                    :data-tip="underlinePart.title"
                    :aria-label="underlinePart.title"
                    :aria-pressed="underlineOn(style, lastUnderline)"
                    :disabled="!formattable"
                    @click="underline(lastUnderline)"
                  >
                    <component :is="underlinePart.icon" />
                  </button>
                  <!-- on a phone the arrow would push Bold out of the row: Double Underline is on Format Cells' Font tab there -->
                  <ToolbarDrop v-if="!narrow" hook="underline-menu" :label="m.underlineMenu" :open="open === 'underline'" :disabled="!formattable" :items="underlineItems" @toggle="toggle('underline')" />
                </span>
              </span>
              <span v-else-if="p.block.group === 'borders'" class="xs-tb-part">
                <span class="xs-tb-split">
                  <button type="button" class="xs-tool" data-tb="border-last" :data-tip="text(lastBorder.title)" :aria-label="text(lastBorder.title)" :disabled="!formattable" @click="border(lastBorder.kind)">
                    <component :is="lastBorder.icon" />
                  </button>
                  <button type="button" class="xs-tool xs-tb-caret" data-tb="borders" :data-tip="m.borders" :aria-label="m.borders" aria-haspopup="menu" :aria-expanded="open === 'borders'" :disabled="!formattable" @click="toggle('borders')">
                    <IconDown />
                  </button>
                  <div v-if="open === 'borders'" class="xs-tb-pop" role="menu" :aria-label="m.borders">
                    <template v-for="(b, i) in BORDERS" :key="b.kind">
                      <div v-if="i === 1" class="xs-tb-sep" role="separator" />
                      <button type="button" class="xs-tb-item" role="menuitem" :data-tb="`border-${b.kind}`" @click="border(b.kind)">
                        <component :is="b.icon" />
                        {{ text(b.title) }}
                      </button>
                    </template>
                  </div>
                </span>
              </span>
              <span v-else-if="p.block.group === 'colors'" class="xs-tb-part">
                <span v-for="c in COLOURS" :key="c.what" class="xs-tb-split">
                  <button type="button" class="xs-tool" :data-tb="`${c.what}-colour`" :data-tip="text(c.title)" :aria-label="text(c.title)" :disabled="!formattable" @click="paint(c.what, last[c.what])">
                    <component :is="c.icon" :colour="last[c.what]" />
                  </button>
                  <button
                    type="button"
                    class="xs-tool xs-tb-caret"
                    :data-tb="`${c.what}-palette`"
                    :data-tip="m.pickColor"
                    :aria-label="m.pickColor"
                    :aria-expanded="open === c.what"
                    :disabled="!formattable"
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
              <span v-else-if="p.block.group === 'verticalAlign'" class="xs-tb-part">
                <button
                  v-for="a in VERTICAL"
                  :key="a.v"
                  type="button"
                  :class="['xs-tool', { 'xs-tool-on': verticalOn(style, a.v) }]"
                  :data-tb="`align-${a.v === 'center' ? 'middle' : a.v}`"
                  :data-tip="text(a.title)"
                  :aria-label="text(a.title)"
                  :aria-pressed="verticalOn(style, a.v)"
                  :disabled="!formattable"
                  @click="restyle(verticalToggle(style, a.v))"
                >
                  <component :is="a.icon" />
                </button>
              </span>
              <span v-else-if="p.block.group === 'orientation'" class="xs-tb-part">
                <span class="xs-tb-split">
                  <ToolbarDrop hook="orientation" :label="m.orientation" :icon="IconOrientation" :open="open === 'orientation'" :disabled="!formattable" :items="orientationItems" @toggle="toggle('orientation')" />
                </span>
              </span>
              <span v-else-if="p.block.group === 'align'" class="xs-tb-part">
                <template v-if="has(p, 'a2')">
                  <button
                    v-for="a in ALIGN"
                    :key="a.h"
                    type="button"
                    :class="['xs-tool', { 'xs-tool-on': style?.alignment?.horizontal === a.h }]"
                    :data-tb="`align-${a.h}`"
                    :data-tip="text(a.title)"
                    :aria-label="text(a.title)"
                    :aria-pressed="style?.alignment?.horizontal === a.h"
                    :disabled="!formattable"
                    @click="apply(over({ style: alignToggle(style, a.h) }))"
                  >
                    <component :is="a.icon" />
                  </button>
                  <button type="button" class="xs-tool" data-tb="indent-down" :data-tip="m.indentDecrease" :aria-label="m.indentDecrease" :disabled="!formattable" @click="restyle(indentStep(style, false))">
                    <IconIndentDecrease />
                  </button>
                  <button type="button" class="xs-tool" data-tb="indent-up" :data-tip="m.indentIncrease" :aria-label="m.indentIncrease" :disabled="!formattable" @click="restyle(indentStep(style, true))">
                    <IconIndentIncrease />
                  </button>
                </template>
                <button
                  v-if="has(p, 'b1')"
                  type="button"
                  :class="['xs-tool', { 'xs-tool-on': !!style?.alignment?.wrap }]"
                  data-tb="wrap"
                  :data-tip="m.wrapText"
                  :aria-label="m.wrapText"
                  :aria-pressed="!!style?.alignment?.wrap"
                  :disabled="!formattable"
                  @click="apply(over({ style: wrapToggle(style) }))"
                >
                  <IconWrap />
                </button>
              </span>
              <span v-else-if="p.block.group === 'merge' && sheet && source.structure" class="xs-tb-part">
                <span class="xs-tb-split">
                  <button
                    type="button"
                    :class="['xs-tool', { 'xs-tool-on': merged }]"
                    data-tb="merge"
                    :data-tip="m.mergeCenter"
                    :aria-label="m.mergeCenter"
                    :aria-pressed="merged"
                    :disabled="!state?.canMerge"
                    @click="merge(merged ? 'unmergeCells' : 'mergeCenter')"
                  >
                    <IconMerge />
                  </button>
                  <button type="button" class="xs-tool xs-tb-caret" data-tb="merge-menu" :data-tip="m.mergeMenu" :aria-label="m.mergeMenu" :aria-expanded="open === 'merge'" :disabled="!state?.canMerge" @click="toggle('merge')">
                    <IconDown />
                  </button>
                  <div v-if="open === 'merge'" class="xs-tb-pop" role="menu" :aria-label="m.mergeMenu">
                    <button v-for="x in MERGES" :key="x" type="button" class="xs-tb-item" role="menuitem" :data-tb="kebab(x)" @click="merge(x)">
                      {{ m[x] }}
                    </button>
                  </div>
                </span>
              </span>
              <span v-else-if="p.block.group === 'number'" class="xs-tb-part">
                <select
                  v-if="has(p, 'a1') && numbers.choices.value.length"
                  class="xs-field xs-field-s xs-tb-select"
                  data-tb="number"
                  :data-tb-category="category?.category"
                  :data-tip="m.numberFormat"
                  :aria-label="m.numberFormat"
                  value=""
                  :disabled="!formattable"
                  @change="pickNumber"
                >
                  <option value="" hidden>{{ category?.name ?? '' }}</option>
                  <option v-for="c in numbers.choices.value" :key="c.key" :value="c.key">{{ c.name }}</option>
                </select>
                <template v-if="has(p, 'a2')">
                  <span class="xs-tb-split">
                    <button type="button" class="xs-tool" data-tb="accounting" :data-tip="m.accountingStyle" :aria-label="m.accountingStyle" :disabled="!formattable" @click="apply(over({ numberFormat: ACCOUNTING_STYLE }))">
                      <IconAccounting />
                    </button>
                    <ToolbarDrop hook="accounting-menu" :label="m.accountingMenu" :open="open === 'accounting'" :disabled="!formattable" :items="accountingItems" @toggle="toggle('accounting')" />
                  </span>
                  <button type="button" class="xs-tool" data-tb="percent" :data-tip="m.percentStyle" :aria-label="m.percentStyle" :disabled="!formattable" @click="apply(over({ numberFormat: PERCENT_STYLE }))">
                    <IconPercent />
                  </button>
                  <button type="button" class="xs-tool" data-tb="comma" :data-tip="m.commaStyle" :aria-label="m.commaStyle" :disabled="!formattable" @click="apply(over({ numberFormat: COMMA_STYLE }))">
                    <IconComma />
                  </button>
                  <button type="button" class="xs-tool" data-tb="decimal-up" :data-tip="m.decimalIncrease" :aria-label="m.decimalIncrease" :disabled="!formattable" @click="stepDecimal(true)">
                    <IconDecimalUp />
                  </button>
                  <button type="button" class="xs-tool" data-tb="decimal-down" :data-tip="m.decimalDecrease" :aria-label="m.decimalDecrease" :disabled="!formattable" @click="stepDecimal(false)">
                    <IconDecimalDown />
                  </button>
                </template>
              </span>
              <span v-else-if="p.block.group === 'formatCells' && (dialogToHost() || sheet)" class="xs-tb-part">
                <button type="button" class="xs-tool" data-tb="format-cells" :data-tip="m.formatCellsOpen" :aria-label="m.formatCellsOpen" :disabled="!formattable" @click="formatCells()">
                  <IconFormatCells />
                </button>
              </span>
              <span v-else-if="p.block.group === 'autoSum' && sheet" class="xs-tb-part">
                <span class="xs-tb-split">
                  <button type="button" class="xs-tool" data-tb="auto-sum" :data-tip="m.autoSum" :aria-label="m.autoSum" :disabled="!state?.canAutoSum" @click="autoSum('sum')">
                    <IconAutoSum />
                  </button>
                  <button type="button" class="xs-tool xs-tb-caret" data-tb="auto-sum-menu" :data-tip="m.autoSumMenu" :aria-label="m.autoSumMenu" :aria-expanded="open === 'autoSum'" :disabled="!state?.canAutoSum" @click="toggle('autoSum')">
                    <IconDown />
                  </button>
                  <div v-if="open === 'autoSum'" class="xs-tb-pop" role="menu" :aria-label="m.autoSumMenu">
                    <button v-for="x in autoSums" :key="x.fn" type="button" class="xs-tb-item" role="menuitem" :data-tb="`auto-sum-${x.fn}`" @click="autoSum(x.fn)">
                      {{ text(x.title) }}
                    </button>
                    <button v-if="source.functionNames" type="button" class="xs-tb-item" role="menuitem" data-tb="auto-sum-more" @click="moreFunctions">
                      {{ m.autoSumMore }}
                    </button>
                  </div>
                </span>
              </span>
              <span v-else-if="p.block.group === 'pivotTable' && sheet && source.structure" :class="['xs-tb-part', { 'xs-tb-large': p.large }]">
                <button type="button" class="xs-tool" data-tb="pivot-table" :data-tip="m.pivotInsert" :aria-label="m.pivotInsert" :disabled="!state?.canInsertPivot" @click="insertPivot">
                  <IconPivotTable />
                  <ToolbarLabel v-if="p.large" :name="m.pivotInsert" />
                </button>
              </span>
              <span v-else-if="p.block.group === 'refreshAll' && sheet && source.refreshConnections" :class="['xs-tb-part', { 'xs-tb-large': p.large }]">
                <span class="xs-tb-split">
                  <button type="button" class="xs-tool" data-tb="refresh-all" :data-tip="m.refreshAll" :data-tip-more="m.refreshAllTip" :aria-label="m.refreshAll" :disabled="!state?.canRefreshAll" @click="refreshAll">
                    <IconRefreshAll />
                  </button>
                  <button type="button" class="xs-tool xs-tb-caret" data-tb="refresh-all-menu" :data-tip="m.refreshAllMenu" :aria-label="m.refreshAllMenu" :aria-expanded="open === 'refreshAll'" :disabled="!state?.canRefreshAll && !state?.canRefreshPivot" @click="toggle('refreshAll')">
                    <ToolbarLabel v-if="p.large" :name="m.refreshAll" arrow />
                    <IconDown v-else />
                  </button>
                  <div v-if="open === 'refreshAll'" class="xs-tb-pop" role="menu" :aria-label="m.refreshAllMenu">
                    <button type="button" class="xs-tb-item" role="menuitem" data-tb="refresh-all-item" :title="m.refreshAll" :disabled="!state?.canRefreshAll" @click="refreshAll">{{ m.refreshAllItem }}</button>
                    <button type="button" class="xs-tb-item" role="menuitem" data-tb="refresh-pivot" :title="m.refreshPivot" :disabled="!state?.canRefreshPivot" @click="refreshPivot">{{ m.menuRefresh }}</button>
                  </div>
                </span>
              </span>
              <span v-else-if="(p.block.group === 'pivotAnalyze' || p.block.group === 'pivotDesign') && sheet && state?.pivot && source.pivotCells" :class="['xs-tb-part', { 'xs-tb-large': p.large }]">
                <ToolbarPivot :group="p.block.group" :large="p.large" :sheet="sheet" :state="state" :source="source" :m="m" />
              </span>
              <span v-else-if="p.block.buttons" class="xs-tb-part">
                <template v-for="x in p.block.buttons" :key="x.id">
                  <!-- With a list, the button is one split button, its part and arrow side by side in one cell, as AutoSum -->
                  <component :is="x.menu ? 'span' : Bare" v-bind="x.menu ? { class: ['xs-tb-split', { 'xs-tb-large': large(g, x) }] } : {}">
                    <button
                      v-if="!x.menu || x.run"
                      type="button"
                      :class="['xs-tool', { 'xs-tool-on': x.pressed, 'xs-tb-wide xs-tb-text': !x.icon, 'xs-tb-large': large(g, x) }]"
                      :data-tb-host="x.id"
                      :data-tip="x.label"
                      :aria-label="x.label"
                      :aria-pressed="x.pressed === undefined ? undefined : x.pressed"
                      :disabled="x.disabled"
                      @click="runHost(x)"
                    >
                      <HostIcon v-if="x.icon || !large(g, x)" :button="x" />
                      <ToolbarLabel v-if="large(g, x) && !x.menu" :name="x.label" />
                    </button>
                    <template v-if="x.menu">
                      <button
                        type="button"
                        :class="['xs-tool', x.run ? 'xs-tb-caret' : 'xs-tb-wide', { 'xs-tool-on': !x.run && x.pressed, 'xs-tb-text': !x.run && !x.icon }]"
                        :data-tb-host-menu="x.id"
                        :data-tip="x.label"
                        :aria-label="x.label"
                        aria-haspopup="menu"
                        :aria-expanded="open === `host:${x.id}`"
                        :disabled="x.disabled"
                        @click="toggle(`host:${x.id}`)"
                      >
                        <HostIcon v-if="!x.run && (x.icon || !large(g, x))" :button="x" />
                        <ToolbarLabel v-if="large(g, x)" :name="x.label" arrow />
                        <IconDown v-else :class="{ 'xs-tb-down': !x.run }" />
                      </button>
                      <div v-if="open === `host:${x.id}`" class="xs-tb-pop" role="menu" :aria-label="x.label">
                        <button v-for="it in x.menu" :key="it.id" type="button" class="xs-tb-item" role="menuitem" :data-tb-host-item="it.id" :disabled="it.disabled" @click="runHost(it)">
                          {{ it.label }}
                        </button>
                      </div>
                    </template>
                  </component>
                </template>
              </span>
            </template>
          </span>
        </span>
      </span>
      <button v-if="part.place === 'row' && moreShown" type="button" :class="['xs-tool', 'xs-tb-more', { 'xs-tool-on': menu }]" data-tb="more" :data-tip="m.toolbarMore" :aria-label="m.toolbarMore" :aria-expanded="menu" @click="menu = !menu">
        <IconMore />
      </button>
    </div>
    <div v-if="tip" ref="tipEl" :class="['xs-tb-tip', { 'xs-tb-tip-more': tip.more }]" role="tooltip" data-tb-tip :style="{ left: `${tip.left}px`, top: `${tip.top}px` }">
      <template v-if="tip.more"><b>{{ tip.text }}</b><span>{{ tip.more }}</span></template>
      <template v-else>{{ tip.text }}</template>
    </div>
  </div>
</template>

<style src="../toolbar.css"></style>
