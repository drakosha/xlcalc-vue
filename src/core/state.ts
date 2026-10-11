// The component's base state: the book's sheets, the active one, its extent, sizes from the windows, and the frame.
import { computed, ref, shallowRef } from 'vue'
import { MAX_COLS, MAX_ROWS } from 'xlcalc/a1'
import type { CellPosition } from 'xlcalc/contract'
import { enMessages, type SheetMessages } from '../messages'
import type { SheetFilter, SheetInfo, SheetSource } from '../source'
import { colPx, DEFAULT_ROW_PT, rowPx } from './geometry'
import { createMemory } from './storage'

/** Properties of `SheetView`. */
export interface SheetViewProps {
  source: SheetSource
  /** Interface texts; a part overrides the English defaults. */
  messages?: Partial<SheetMessages>
  /** Key of the browser's memory for the freeze choice and widths; without it the memory lives with the component. */
  storageKey?: string
  /** A place to show: `Sheet!B12:D15`, `B7`, `Sheet`; a new object is a new jump. */
  focus?: { ref: string } | null
  /** The host's page address for a place; without it there is no "Link to the cell" button. */
  cellUrl?: (ref: string) => string
  readOnly?: boolean
  /** The formula bar over the grid, on by default. */
  formulaBar?: boolean
  /** The book's locale for numbers the component writes itself: a number's edit text and selection totals. */
  locale?: string
}

/** Why the book cannot be shown: the source's error and its text. */
export interface SheetFailure {
  error: unknown
  text: string
}

/** Rows and columns past the data an editable sheet shows, as Excel's grid goes on. */
const SPARE_ROWS = 50
const SPARE_COLS = 5
const MIN_ROWS = 100
const MIN_COLS = 26

export function createState(props: SheetViewProps) {
  const source = () => props.source
  const messages = computed<SheetMessages>(() => ({ ...enMessages, ...props.messages }))
  const locale = computed(() => props.locale ?? 'en-US')
  const memory = createMemory(() => props.storageKey)
  const editable = computed(() => !props.readOnly && typeof props.source.edit === 'function')

  const sheets = shallowRef<SheetInfo[]>([])
  const active = ref(-1)
  const sheet = computed<SheetInfo | null>(() => sheets.value[active.value] ?? null)
  const sheetName = computed(() => sheet.value?.name ?? '')
  /** Data reach grown by edits since the sheets were read: the source's extent is from before them. */
  const grown = ref({ rows: 0, cols: 0 })
  /** Reach of jumps and keys past the data on an editable sheet. */
  const reach = ref({ rows: 0, cols: 0 })
  const dataRows = computed(() => Math.max(sheet.value?.maxRow ?? 0, grown.value.rows))
  const dataCols = computed(() => Math.max(sheet.value?.maxCol ?? 0, grown.value.cols))
  const failure = shallowRef<SheetFailure | null>(null)
  const loading = ref(false)
  const scroller = ref<HTMLElement | null>(null)
  const scrollTop = ref(0)
  const scrollLeft = ref(0)
  const viewH = ref(600)
  const viewW = ref(800)
  /** The display filter in force: the sheet's rows in display order and the way back; null without one. */
  const visible = shallowRef<{ key: string; spec: SheetFilter; rows: Int32Array; index: Map<number, number> } | null>(null)
  /** A filter's list of rows is on its way: windows wait for it, their rows would be another's. */
  const filterWait = ref(false)
  /** Display index (`d`, the grid's geometry and windows) to the sheet's row (`r`, numbers, addresses, selection) and back. */
  const rowOf = (d: number): number => (visible.value ? (visible.value.rows[d - 1] ?? d) : d)
  const indexOf = (r: number): number | undefined => (visible.value ? visible.value.index.get(r) : r)
  /** Rows with data in display order: the filter's list, else the data rows. */
  const dataShown = computed(() => (visible.value ? visible.value.rows.length : dataRows.value))
  /** Rows the grid shows: a filter's list as is; an editable sheet goes on past the data, up to Excel's last row. */
  const rowCount = computed(() => {
    if (visible.value) return visible.value.rows.length
    return editable.value ? Math.min(MAX_ROWS, Math.max(dataRows.value + SPARE_ROWS, MIN_ROWS, reach.value.rows)) : dataRows.value
  })
  /** The last sheet row a selection of whole columns reaches. */
  const lastRow = computed(() => (visible.value ? Math.max(1, dataRows.value) : Math.max(1, rowCount.value)))
  /** An editable sheet fills the frame with columns at least, up to Excel's last column. */
  const colCount = computed(() =>
    editable.value ? Math.min(MAX_COLS, Math.max(dataCols.value + SPARE_COLS, MIN_COLS, reach.value.cols, Math.ceil(viewW.value / colPx(null)) + 1)) : dataCols.value,
  )

  const bookWidths = shallowRef(new Map<number, number>())
  const userWidths = ref<Record<number, number>>({})
  const heights = shallowRef(new Map<number, number>())
  const defaultRowPx = ref(rowPx(DEFAULT_ROW_PT))

  /** The active cell and the selection's other corner. */
  const selected = ref<CellPosition | null>(null)
  const extent = ref<CellPosition | null>(null)

  /** Ways back into modules made later: dropping the filter when a jump needs a row it hides. */
  const hooks = { unfilter: (_row: number): void => {} }

  /** Grow the data reach to a cell that got a value. */
  function grow(row: number, col: number): void {
    if (row > dataRows.value || col > dataCols.value)
      grown.value = { rows: Math.max(grown.value.rows, row), cols: Math.max(grown.value.cols, col) }
  }
  /** Let an editable sheet show a far cell: a jump or a key went past the shown grid. */
  function stretch(row: number, col: number): void {
    if (!editable.value) return
    if (row + SPARE_ROWS > rowCount.value || col + SPARE_COLS > colCount.value)
      reach.value = { rows: Math.max(reach.value.rows, row + SPARE_ROWS), cols: Math.max(reach.value.cols, col + SPARE_COLS) }
  }

  return {
    props,
    source,
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
    hooks,
    grow,
    stretch,
  }
}

export type SheetState = ReturnType<typeof createState>
