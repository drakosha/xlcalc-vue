// The component's base state: the workbook's sheets, the active one, its extent, sizes from the windows, and the viewport.
import { computed, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { MAX_COLS, MAX_ROWS } from 'xlcalc/a1'
import type { CellPosition } from 'xlcalc/contract'
import { DEFAULT_ROW_HEIGHT } from 'xlcalc/text'
import { errorText } from '../failure'
import { useMessages } from '../messages'
import type { SheetMessages } from '../contract/messages'
import type { SheetFilter, SheetInfo, SheetSource } from '../contract/source'
import { colPx, rowPx } from './geometry'
import type { ListSort, SheetFailure, SheetViewProps } from '../contract/view'
import { createMemory } from './storage'

/** Rows and columns beyond the data that an editable sheet displays, as Excel's grid continues. */
const SPARE_ROWS = 50
const SPARE_COLS = 5
const MIN_ROWS = 100
const MIN_COLS = 26

/** The source's methods by name, optional ones too. */
export type SourceMethod = { [K in keyof SheetSource]-?: NonNullable<SheetSource[K]> extends (...args: never[]) => unknown ? K : never }[keyof SheetSource]

export function createState(props: SheetViewProps) {
  const source = () => props.source
  /** A source method bound to the source (a class's methods need their `this`), or null when the source lacks it. */
  function sourceMethod<K extends SourceMethod>(name: K): NonNullable<SheetSource[K]> | null {
    const src = props.source
    const fn = src[name] as ((...args: unknown[]) => unknown) | undefined
    return typeof fn === 'function' ? (fn.bind(src) as NonNullable<SheetSource[K]>) : null
  }
  let generation = 0
  /** A check for a call's late answer: true once a new source or the sheet's unmount has replaced the workbook the call was made in. */
  function outdated(): () => boolean {
    const at = generation
    return () => at !== generation
  }
  /** A new source or the unmount: the calls made so far drop their answers. */
  const endWorkbook = () => void generation++
  const messages = useMessages(() => props.messages)
  const locale = computed(() => props.locale ?? 'en-US')
  const memory = createMemory(() => props.storageKey)
  const editable = computed(() => !props.readOnly && typeof props.source.edit === 'function')

  const sheets = shallowRef<SheetInfo[]>([])
  const active = ref(-1)
  const sheet = computed<SheetInfo | null>(() => sheets.value[active.value] ?? null)
  const sheetName = computed(() => sheet.value?.name ?? '')
  /** The data extent grown by edits since the sheets were read: the source's extent predates them. */
  const grown = ref({ rows: 0, cols: 0 })
  /** The extent reached by jumps and keys beyond the data on an editable sheet. */
  const reach = ref({ rows: 0, cols: 0 })
  const dataRows = computed(() => Math.max(sheet.value?.maxRow ?? 0, grown.value.rows))
  const dataCols = computed(() => Math.max(sheet.value?.maxCol ?? 0, grown.value.cols))
  const failure = shallowRef<SheetFailure | null>(null)
  /** Why the workbook cannot be displayed, from a source's error: its message, otherwise the interface's. */
  const failureOf = (error: unknown): SheetFailure => ({ error, text: errorText(error) || messages.value.loadFailed })
  const loading = ref(false)
  const scroller = ref<HTMLElement | null>(null)
  const scrollTop = ref(0)
  const scrollLeft = ref(0)
  const viewH = ref(600)
  const viewW = ref(800)
  /** The display filter in effect: the sheet's rows in display order and the reverse map; null without one. */
  const visible = shallowRef<{ key: string; spec: SheetFilter; rows: Int32Array; index: Int32Array } | null>(null)
  /** A filter's row list is loading: windows wait for it, since their rows would belong to another list. */
  const filterWait = ref(false)
  /** Display index (`d`, the grid's geometry and windows) to sheet row (`r`, row numbers, addresses, selection) and back. */
  const rowOf = (d: number): number => (visible.value ? (visible.value.rows[d - 1] ?? d) : d)
  const indexOf = (r: number): number | undefined => (visible.value ? visible.value.index[r] || undefined : r)
  /** Rows with data in display order: the filter's list, otherwise the data rows. */
  const dataShown = computed(() => (visible.value ? visible.value.rows.length : dataRows.value))
  /** Rows the grid displays: a filter's list as is; an editable sheet continues beyond the data, up to Excel's last row. */
  const rowCount = computed(() => {
    if (visible.value) return visible.value.rows.length
    return editable.value ? Math.min(MAX_ROWS, Math.max(dataRows.value + SPARE_ROWS, MIN_ROWS, reach.value.rows)) : dataRows.value
  })
  /** The last sheet row a selection of entire columns reaches. */
  const lastRow = computed(() => (visible.value ? Math.max(1, dataRows.value) : Math.max(1, rowCount.value)))
  /** An editable sheet fills the viewport with columns at least, up to Excel's last column. */
  const colCount = computed(() =>
    editable.value ? Math.min(MAX_COLS, Math.max(dataCols.value + SPARE_COLS, MIN_COLS, reach.value.cols, Math.ceil(viewW.value / colPx(null)) + 1)) : dataCols.value,
  )

  /** Column widths as the workbook stores them (null: the default); the grid draws them in the library's geometry. */
  const bookWidths = shallowRef(new Map<number, number | null>())
  const userWidths = shallowRef<Record<number, number>>({})
  const heights = shallowRef(new Map<number, number>())
  const defaultRowPx = ref(rowPx(DEFAULT_ROW_HEIGHT))

  /** The active cell and the selection's other corner. */
  const selected = ref<CellPosition | null>(null)
  const extent = ref<CellPosition | null>(null)
  /** The selection's corner opposite `extent` once Enter or Tab moved the active cell inside it; null: the active cell is that corner. */
  const anchor = ref<CellPosition | null>(null)

  /** Hooks into modules created later: clearing the filter when a jump needs a row it hides; a change of cells coming from the sheet (Excel's copy mode ends). */
  const hooks = { unfilter: (_row: number): void => {}, changing: (): void => {} }

  /** The grid takes focus back without scrolling the page: after an editor, a popover, a dialog or a command. */
  const focusGrid = (): void => scroller.value?.focus({ preventScroll: true })
  /** Scrolls the viewport and records the scroll in the state at once: what is drawn follows without waiting for the scroll event. */
  function scrollGridTo(top: number, left: number): void {
    scroller.value?.scrollTo({ top, left })
    scrollTop.value = top
    scrollLeft.value = left
  }

  /** Grows the data extent to a cell that got a value. */
  function grow(row: number, col: number): void {
    if (row > dataRows.value || col > dataCols.value)
      grown.value = { rows: Math.max(grown.value.rows, row), cols: Math.max(grown.value.cols, col) }
  }
  /** Lets an editable sheet display a distant cell: a jump or a key went beyond the displayed grid. */
  function stretch(row: number, col: number): void {
    if (!editable.value) return
    if (row + SPARE_ROWS > rowCount.value || col + SPARE_COLS > colCount.value)
      reach.value = { rows: Math.max(reach.value.rows, row + SPARE_ROWS), cols: Math.max(reach.value.cols, col + SPARE_COLS) }
  }

  return {
    props,
    source,
    sourceMethod,
    outdated,
    endWorkbook,
    messages,
    locale,
    memory,
    editable,
    sheets,
    active,
    sheet,
    sheetName,
    grown,
    reach,
    dataRows,
    dataCols,
    visible,
    filterWait,
    rowOf,
    indexOf,
    dataShown,
    rowCount,
    lastRow,
    colCount,
    failure,
    failureOf,
    loading,
    scroller,
    scrollTop,
    scrollLeft,
    viewH,
    viewW,
    bookWidths,
    userWidths,
    heights,
    defaultRowPx,
    selected,
    extent,
    anchor,
    hooks,
    focusGrid,
    scrollGridTo,
    grow,
    stretch,
  }
}

export type SheetState = ReturnType<typeof createState>

/** Signals for reads of the displayed sheet (a series of windows): each new one aborts the previous, and so do another sheet and unmounting the sheet. */
export function useSheetReads(s: SheetState): () => AbortSignal {
  let ctrl: AbortController | null = null
  const stop = () => ctrl?.abort()
  watch(s.sheetName, stop)
  onScopeDispose(stop)
  return () => {
    stop()
    return (ctrl = new AbortController()).signal
  }
}
