// Excel's dialogs and commands over a PivotTable: Value Field Settings, Field Settings, PivotTable Options, Grouping,
// Insert Calculated Field and Item, Ungroup and the Design tab, each OK one structure operation (one step of undo) through
// the source's `structure`; the PivotTables of the displayed sheet and the cell under the active cell, read from the
// source's `pivotTable` and `pivotCells` for the cell menu, the filter buttons and the dialogs.
import { computed, markRaw, ref, shallowRef, watch } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import { renamedValueField } from 'xlcalc/text'
import type { Locale, NumberFormat, PivotDataField, PivotFunction, PivotGroupBy, PivotGrouping, PivotShowAs, PivotTablePatch } from 'xlcalc/contract'
import type { PivotOptionsTab } from '../contract/messages'
import type { EditOutcome, SheetOp, SheetPivotCell } from '../contract/source'
import type { PivotDesignChange } from '../contract/view'
import { loadEntry } from '../entry'
import { axisFields, cellField, designPatch, grandOf, grandPatch, ungroupedField, fieldKind, fieldName, groupedField, isoOfSerial, PIVOT_FUNCTIONS, serialOfIso, withFunction, withValue, type PivotTarget } from '../pivot-model'
import type { Commands } from './commands'
import type { Pivots } from './pivots'
import type { Protection } from './protection'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'

/** The calculations of Show Values As that take a base field, and those of them that take a base item too. */
export const BASE_FIELD = new Set<PivotShowAs>(['percent', 'percentOfParent', 'difference', 'percentDiff', 'runTotal', 'percentOfRunningTotal', 'rankAscending', 'rankDescending'])
export const BASE_ITEM = new Set<PivotShowAs>(['percent', 'difference', 'percentDiff'])
/** Show Values As's list in Excel's order. */
export const SHOW_AS: readonly PivotShowAs[] = ['normal', 'percentOfTotal', 'percentOfColumn', 'percentOfRow', 'percent', 'percentOfParentRow', 'percentOfParentColumn', 'percentOfParent', 'difference', 'percentDiff', 'runTotal', 'percentOfRunningTotal', 'rankAscending', 'rankDescending', 'index']
/** The Grouping dialog's By list, finest first, as Excel lists it. */
export const PERIODS: readonly Exclude<PivotGroupBy, 'numbers'>[] = ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']
/** The most selected cells whose items Group and Hide Selected Items read. */
const MAX_PICKED = 2000

export interface ValueDraft {
  name: string
  fn: PivotFunction
  showAs: PivotShowAs
  baseField: number
  baseItem: number | 'previous' | 'next'
  numberFormat?: NumberFormat
}
export interface FieldDraft {
  name: string
  subtotals: 'auto' | 'none' | 'custom'
  custom: PivotFunction[]
  includeNewItems: boolean
  form: 'outline' | 'tabular'
  compact: boolean
  subtotalTop: boolean
  repeatLabels: boolean
  blankLine: boolean
  showEmpty: boolean
  pageBreak: boolean
}
export interface OptionsDraft {
  name: string
  mergeLabels: boolean
  indent: string
  pagesAcrossFirst: boolean
  pagesWrap: string
  errorsOn: boolean
  errors: string
  emptyOn: boolean
  empty: string
  autofitColumns: boolean
  preserveFormatting: boolean
  grandRows: boolean
  grandColumns: boolean
  subtotalHiddenItems: boolean
  multipleFieldFilters: boolean
  customListSort: boolean
  showDrill: boolean
  fieldHeaders: boolean
  classicLayout: boolean
  valuesRow: boolean
  enableDrill: boolean
  altTitle: string
  altDescription: string
}
export interface GroupDraft {
  dates: boolean
  autoStart: boolean
  start: string
  autoEnd: boolean
  end: string
  by: string
  periods: Exclude<PivotGroupBy, 'numbers'>[]
  days: string
  error: string
}
export interface CalcDraft {
  name: string
  formula: string
  pick: number
  item: number
}

/** The open dialog over a PivotTable, with what it edits. */
export type PivotDialogState =
  | { kind: 'value'; t: PivotTarget; index: number; tab: 'summarize' | 'showAs'; draft: ValueDraft; format: boolean }
  | { kind: 'numberFormat'; t: PivotTarget; index: number; format?: NumberFormat }
  | { kind: 'field'; t: PivotTarget; field: number; tab: 'subtotals' | 'layout'; draft: FieldDraft }
  | { kind: 'options'; t: PivotTarget; tab: PivotOptionsTab; draft: OptionsDraft; start: OptionsDraft }
  | { kind: 'group'; t: PivotTarget; field: number; draft: GroupDraft }
  | { kind: 'calcField'; t: PivotTarget; draft: CalcDraft }
  | { kind: 'calcItem'; t: PivotTarget; field: number; draft: CalcDraft }

export interface PivotDialogsDeps {
  /** The workbook's locale as xlcalc's text functions take it, an object once its module is loaded. */
  locale: () => string | Locale
  pivots: Pivots
  protection: Protection
  sel: Selection
}

export function usePivotDialogs(s: SheetState, cmd: Commands, deps: PivotDialogsDeps) {
  const readViews = useSheetReads(s)
  const readHere = useSheetReads(s)
  const read = useSheetReads(s)
  const m = () => s.messages.value
  const can = computed(() => !!deps.pivots.active.value && cmd.can('structure') && !!s.sourceMethod('pivotTable') && !!s.sourceMethod('pivotCells') && deps.protection.can('pivotTables'))
  const state = ref<PivotDialogState | null>(null)
  /** The PivotTable stays a plain object: an operation built from it goes to a worker, which cannot clone a reactive proxy. */
  const open = (next: PivotDialogState) => (state.value = { ...next, t: markRaw(next.t) })
  const close = () => {
    state.value = null
    s.focusGrid()
  }

  // ── the PivotTables read from the source ──
  /** A PivotTable's definition and cells, with what `cells` are; null without the source's methods or the PivotTable. */
  async function target(sheet: string, pivot: string, cells?: string[], signal = read()): Promise<{ t: PivotTarget; parts: (SheetPivotCell | null)[] } | null> {
    const src = s.source()
    if (!src.pivotTable || !src.pivotCells) return null
    try {
      const [def, got] = await Promise.all([src.pivotTable(sheet, pivot, signal), src.pivotCells(sheet, pivot, cells, signal)])
      if (!def || !got || signal.aborted) return null
      return { t: { sheet, pivot, def, cells: got }, parts: got.cells ?? [] }
    } catch {
      return null
    }
  }
  /** The PivotTables of the displayed sheet with their buttons, read again as the sheets are. */
  const views = shallowRef<PivotTarget[]>([])
  watch(
    () => [s.sheetName.value, s.sheet.value?.pivots] as const,
    async ([sheet, pivots]) => {
      if (!pivots?.length || !s.sourceMethod('pivotCells')) return void (views.value = [])
      const signal = readViews()
      const got = await Promise.all(pivots.map((p) => target(sheet, p.name, undefined, signal)))
      if (!signal.aborted) views.value = got.flatMap((x) => (x ? [x.t] : []))
    },
    { immediate: true },
  )
  /** The PivotTable under the active cell and what that cell is, for the cell menu; null outside one or while it is read. */
  const here = shallowRef<{ t: PivotTarget; part: SheetPivotCell; cell: string } | null>(null)
  watch(
    () => [deps.pivots.active.value, s.selected.value, views.value] as const,
    async ([p, a]) => {
      here.value = null
      if (!p || !a) return
      const cell = cellAddress(a.row, a.col)
      const signal = readHere()
      const got = await target(s.sheetName.value, p.name, [cell], signal)
      if (got?.parts[0] && !signal.aborted) here.value = { t: got.t, part: got.parts[0], cell }
    },
    { immediate: true },
  )
  /** What the selected cells are in the PivotTable under the active cell (Group, Hide Selected Items). */
  async function picked(): Promise<{ t: PivotTarget; parts: SheetPivotCell[] } | null> {
    const p = deps.pivots.active.value
    const r = deps.sel.fullRect.value
    if (!p || !r || !can.value) return null
    const cells: string[] = []
    for (let row = r.top; row <= r.bottom && cells.length < MAX_PICKED; row++) for (let col = r.left; col <= r.right && cells.length < MAX_PICKED; col++) cells.push(cellAddress(row, col))
    const got = await target(s.sheetName.value, p.name, cells)
    return got && { t: got.t, parts: got.parts.filter((x): x is SheetPivotCell => !!x) }
  }
  /** The PivotTable under the active cell with that cell's part, read now. */
  async function active(): Promise<{ t: PivotTarget; part: SheetPivotCell | null } | null> {
    const p = deps.pivots.active.value
    const a = s.selected.value
    if (!p || !a || !can.value) return null
    const got = await target(s.sheetName.value, p.name, [cellAddress(a.row, a.col)])
    return got && { t: got.t, part: got.parts[0] ?? null }
  }

  // ── sending ──
  const send = (op: SheetOp): Promise<EditOutcome | null> => cmd.structure(op)
  const edit = (t: PivotTarget, set: PivotTablePatch) => send({ op: 'editPivotTable', sheet: t.sheet, pivot: t.pivot, set })

  // ── Value Field Settings ──
  function valueDraft(t: PivotTarget, index: number): ValueDraft | null {
    const v = t.def.table.values?.[index]
    if (!v) return null
    const base = axisFields(t, 'rows')[0] ?? axisFields(t, 'columns')[0] ?? 0
    return { name: v.name ?? '', fn: v.function ?? 'sum', showAs: v.showAs ?? 'normal', baseField: v.baseField ?? base, baseItem: v.baseItem ?? 'previous', ...(v.numberFormat !== undefined ? { numberFormat: v.numberFormat } : {}) }
  }
  function openValue(t: PivotTarget, index: number, tab: 'summarize' | 'showAs' = 'summarize'): boolean {
    const d = valueDraft(t, index)
    if (!d) return false
    open({ kind: 'value', t, index, tab, draft: d, format: false })
    return true
  }
  /** The source field's name of a value field, which its default name goes by. */
  const sourceName = (t: PivotTarget, index: number) => t.def.cache.fields[t.def.table.values?.[index]?.field ?? -1]?.name ?? ''
  /** Another function picked in the list: Custom Name changes at once as Excel's does (xlcalc's `renamedValueField`). */
  function pickFunction(fn: PivotFunction): void {
    const d = state.value
    if (d?.kind !== 'value' || d.draft.fn === fn) return
    d.draft.name = renamedValueField(d.draft.name, d.draft.fn, fn, sourceName(d.t, d.index), s.locale.value)
    d.draft.fn = fn
  }
  /** Another base field: its base item starts again at the list's first, (previous). */
  function pickBaseField(field: number): void {
    const d = state.value
    if (d?.kind !== 'value') return
    d.draft.baseField = field
    d.draft.baseItem = 'previous'
  }
  /** OK: the value field as the dialog sets it, named as Custom Name holds it (xlcalc's name for an empty box). */
  async function applyValue(): Promise<void> {
    const d = state.value
    if (d?.kind !== 'value') return
    const v = d.t.def.table.values![d.index]!
    const x = d.draft
    const base = withFunction(v, x.fn, x.name.trim() || undefined)
    const { showAs: _s, baseField: _f, baseItem: _i, numberFormat: _n, ...rest } = base
    const next: PivotDataField = {
      ...rest,
      ...(x.showAs !== 'normal' ? { showAs: x.showAs } : {}),
      ...(BASE_FIELD.has(x.showAs) ? { baseField: x.baseField } : {}),
      ...(BASE_ITEM.has(x.showAs) ? { baseItem: x.baseItem } : {}),
      ...(x.numberFormat !== undefined ? { numberFormat: x.numberFormat } : {}),
    }
    close()
    await edit(d.t, withValue(d.t, d.index, next))
  }
  /** Summarize Values By and Show Values As from the cell menu: one value field changed at once. */
  async function setFunction(t: PivotTarget, index: number, fn: PivotFunction): Promise<void> {
    const v = t.def.table.values?.[index]
    const from = v?.function ?? 'sum'
    // without a name xlcalc's editPivotTable renames the value field as Excel does
    if (v && from !== fn) await edit(t, withValue(t, index, withFunction(v, fn)))
  }
  async function setShowAs(t: PivotTarget, index: number, calc: PivotShowAs): Promise<void> {
    if (BASE_FIELD.has(calc)) {
      openValue(t, index, 'showAs')
      const d = state.value
      if (d?.kind === 'value') d.draft.showAs = calc
      return
    }
    const v = t.def.table.values?.[index]
    if (!v) return
    const { showAs: _s, baseField: _f, baseItem: _i, ...rest } = v
    await edit(t, withValue(t, index, { ...rest, ...(calc !== 'normal' ? { showAs: calc } : {}) }))
  }
  /** Number Format…: Format Cells' Number tab over a value field's format, from the cell menu or Value Field Settings. */
  function openNumberFormat(t: PivotTarget, index: number): void {
    const d = state.value
    const v = t.def.table.values?.[index]
    if (d?.kind === 'value') d.format = true
    else open({ kind: 'numberFormat', t, index, ...(v?.numberFormat !== undefined ? { format: v.numberFormat } : {}) })
  }
  /** OK of Number Format…: the format chosen or typed, a built-in one by its id; null when OK left the format as it was. */
  async function applyNumberFormat(format: NumberFormat | null): Promise<void> {
    const d = state.value
    if (d?.kind === 'value') {
      if (format !== null) d.draft.numberFormat = format
      d.format = false
      return
    }
    if (d?.kind !== 'numberFormat') return
    const v = d.t.def.table.values?.[d.index]
    close()
    if (v && format !== null) await edit(d.t, withValue(d.t, d.index, { ...v, numberFormat: format }))
  }

  // ── Field Settings ──
  function openField(t: PivotTarget, field: number): boolean {
    const tb = t.def.table
    const f = tb.fields[field]
    if (!t.def.cache.fields[field]) return false
    const layout = f?.layout ?? tb.layout ?? 'compact'
    const subs = f?.subtotals
    open({
      kind: 'field',
      t,
      field,
      tab: 'subtotals',
      draft: {
        name: fieldName(t, field),
        subtotals: subs === undefined ? 'auto' : subs.length ? 'custom' : 'none',
        custom: subs?.length ? [...subs] : [],
        includeNewItems: !!f?.includeNewItems,
        form: layout === 'tabular' ? 'tabular' : 'outline',
        compact: layout === 'compact',
        subtotalTop: f?.subtotalTop !== false,
        repeatLabels: f?.repeatLabels ?? !!tb.repeatLabels,
        blankLine: !!f?.blankLine,
        showEmpty: !!f?.showEmpty,
        pageBreak: !!f?.pageBreak,
      },
    })
    return true
  }
  async function applyField(): Promise<void> {
    const d = state.value
    if (d?.kind !== 'field') return
    const x = d.draft
    const tb = d.t.def.table
    const own = d.t.def.cache.fields[d.field]!.name
    const layout = x.form === 'tabular' ? 'tabular' : x.compact ? 'compact' : 'outline'
    const custom = PIVOT_FUNCTIONS.filter((fn) => x.custom.includes(fn))
    close()
    await edit(d.t, {
      fields: [
        {
          field: d.field,
          name: x.name.trim() && x.name.trim() !== own ? x.name.trim() : null,
          subtotals: x.subtotals === 'auto' ? null : x.subtotals === 'none' ? [] : custom,
          includeNewItems: x.includeNewItems || null,
          layout: layout === (tb.layout ?? 'compact') ? null : layout,
          subtotalTop: x.subtotalTop ? null : false,
          repeatLabels: x.repeatLabels === !!tb.repeatLabels ? null : x.repeatLabels,
          blankLine: x.blankLine || null,
          showEmpty: x.showEmpty || null,
          pageBreak: x.pageBreak || null,
        },
      ],
    })
  }

  // ── PivotTable Options ──
  function optionsDraft(t: PivotTarget): OptionsDraft {
    const tb = t.def.table
    return {
      name: t.pivot,
      mergeLabels: !!tb.mergeLabels,
      indent: String(tb.indent ?? 1),
      pagesAcrossFirst: !!tb.pagesAcrossFirst,
      pagesWrap: String(tb.pagesWrap ?? 0),
      errorsOn: !!tb.errors?.show,
      errors: tb.errors?.text ?? '',
      emptyOn: tb.empty?.show !== false,
      empty: tb.empty?.text ?? '',
      autofitColumns: tb.autofitColumns !== false,
      preserveFormatting: tb.preserveFormatting !== false,
      grandRows: grandOf(t).rows,
      grandColumns: grandOf(t).columns,
      subtotalHiddenItems: !!tb.subtotalHiddenItems,
      multipleFieldFilters: !!tb.multipleFieldFilters,
      customListSort: tb.customListSort !== false,
      showDrill: tb.showDrill !== false,
      fieldHeaders: tb.fieldHeaders !== false,
      classicLayout: !!tb.classicLayout,
      valuesRow: !!tb.valuesRow,
      enableDrill: tb.enableDrill !== false,
      altTitle: tb.altText?.title ?? '',
      altDescription: tb.altText?.description ?? '',
    }
  }
  function openOptions(t: PivotTarget): void {
    const d = optionsDraft(t)
    open({ kind: 'options', t, tab: 'layout', draft: d, start: { ...d } })
  }
  /** The options' boxes that are not whole numbers in Excel's ranges (indent 0 to 127, filter fields per column 0 to 255). */
  const wholeIn = (text: string, max: number) => /^\d+$/.test(text.trim()) && Number(text) <= max
  const optionsValid = (x: OptionsDraft) => wholeIn(x.indent, 127) && wholeIn(x.pagesWrap, 255) && !!x.name.trim()
  /** OK: one edit of what changed, and Rename when the name changed (xlcalc's `renamePivotTable`, a step of its own). */
  async function applyOptions(): Promise<void> {
    const d = state.value
    if (d?.kind !== 'options' || !optionsValid(d.draft)) return
    const x = d.draft
    const a = d.start
    const set: PivotTablePatch = {}
    const flag = <K extends keyof OptionsDraft>(k: K, on: (v: OptionsDraft[K]) => void) => x[k] !== a[k] && on(x[k])
    flag('mergeLabels', (v) => (set.mergeLabels = v || null))
    flag('indent', (v) => (set.indent = Number(v)))
    flag('pagesAcrossFirst', (v) => (set.pagesAcrossFirst = v || null))
    flag('pagesWrap', (v) => (set.pagesWrap = Number(v) || null))
    if (x.errorsOn !== a.errorsOn || x.errors !== a.errors) set.errors = { show: x.errorsOn, ...(x.errors ? { text: x.errors } : {}) }
    if (x.emptyOn !== a.emptyOn || x.empty !== a.empty) set.empty = { show: x.emptyOn, ...(x.empty ? { text: x.empty } : {}) }
    flag('autofitColumns', (v) => (set.autofitColumns = v ? null : false))
    flag('preserveFormatting', (v) => (set.preserveFormatting = v ? null : false))
    if (x.grandRows !== a.grandRows || x.grandColumns !== a.grandColumns) set.grandTotals = grandPatch(x.grandRows, x.grandColumns)
    flag('subtotalHiddenItems', (v) => (set.subtotalHiddenItems = v || null))
    flag('multipleFieldFilters', (v) => (set.multipleFieldFilters = v || null))
    flag('customListSort', (v) => (set.customListSort = v ? null : false))
    flag('showDrill', (v) => (set.showDrill = v ? null : false))
    flag('fieldHeaders', (v) => (set.fieldHeaders = v ? null : false))
    flag('classicLayout', (v) => (set.classicLayout = v || null))
    flag('valuesRow', (v) => (set.valuesRow = v || null))
    flag('enableDrill', (v) => (set.enableDrill = v ? null : false))
    if (x.altTitle !== a.altTitle || x.altDescription !== a.altDescription) set.altText = x.altTitle || x.altDescription ? { ...(x.altTitle ? { title: x.altTitle } : {}), ...(x.altDescription ? { description: x.altDescription } : {}) } : null
    close()
    const t = d.t
    const name = x.name.trim()
    if (Object.keys(set).length) {
      const out = await edit(t, set)
      if (!out?.ok) return
    }
    if (name !== t.pivot) await send({ op: 'renamePivotTable', sheet: t.sheet, pivot: t.pivot, name })
  }

  // ── Grouping ──
  /** Opens Grouping over a field of dates or numbers, with its grouping or Excel's automatic start and end. */
  async function openGroup(t: PivotTarget, at: number): Promise<boolean> {
    const field = groupedField(t, at)
    const kind = fieldKind(t, field)
    if (kind === 'text') return false
    const cf = t.def.cache.fields[field]!
    const range = cf.group?.range ?? t.def.cache.fields.find((x) => x.group?.base === field && x.group.range)?.group?.range
    const facts = cf.facts
    const entry = await loadEntry()
    const locale = deps.locale()
    const dates = kind === 'dates'
    const text = (v: number | { date: string } | undefined, plusDay = false): string => {
      if (v === undefined) return ''
      if (typeof v === 'number') return dates ? entry.entryText(v + (plusDay ? 1 : 0), { locale, numberFormat: 14 }) : entry.entryText(v, { locale })
      const serial = serialOfIso(v.date)
      // a field of whole days ends a day after its last date, as Excel takes it (xlcalc's docs/pivot-tables.md, Grouping)
      return entry.entryText(serial + (plusDay && Number.isInteger(serial) ? 1 : 0), { locale, numberFormat: 14 })
    }
    const periods = PERIODS.filter((p) => p === range?.by || t.def.cache.fields.some((x) => x.group?.base === field && x.group.range?.by === p))
    open({
      kind: 'group',
      t,
      field,
      draft: {
        dates,
        autoStart: range?.autoStart !== false,
        start: text(range?.start ?? facts?.min),
        autoEnd: range?.autoEnd !== false,
        end: text(range?.end ?? facts?.max, range?.end === undefined && dates),
        by: entry.entryText(range?.interval ?? (dates ? 1 : 10), { locale }),
        periods: periods.length ? periods : dates ? ['months'] : [],
        days: entry.entryText(range?.by === 'days' ? (range.interval ?? 1) : 1, { locale }),
        error: '',
      },
    })
    return true
  }
  /** OK: the grouping as xlcalc's `groupPivotField`; a box that does not read as a number or a date is Excel's alert. */
  async function applyGroup(): Promise<void> {
    const d = state.value
    if (d?.kind !== 'group') return
    const x = d.draft
    const entry = await loadEntry()
    const locale = deps.locale()
    const num = (text: string) => {
      const got = entry.parseEntry(text, { locale })
      return got.ok && typeof got.input?.value === 'number' ? got.input.value : null
    }
    const start = x.autoStart ? undefined : num(x.start)
    const end = x.autoEnd ? undefined : num(x.end)
    if (start === null || end === null) return void (x.error = x.dates ? m().pivotBadDate : m().pivotBadNumber)
    let group: PivotGrouping
    if (x.dates) {
      if (!x.periods.length) return
      const daysOnly = x.periods.length === 1 && x.periods[0] === 'days'
      const days = daysOnly ? num(x.days) : null
      if (daysOnly && (days === null || days < 1 || !Number.isInteger(days))) return void (x.error = m().pivotBadNumber)
      group = { by: 'dates', periods: [...x.periods], ...(days && days > 1 ? { days } : {}), ...(start !== undefined ? { start: isoOfSerial(start) } : {}), ...(end !== undefined ? { end: isoOfSerial(end) } : {}) }
    } else {
      const by = num(x.by)
      if (by === null || by <= 0) return void (x.error = m().pivotBadNumber)
      group = { by: 'numbers', interval: by, ...(start !== undefined ? { start } : {}), ...(end !== undefined ? { end } : {}) }
    }
    close()
    await send({ op: 'groupPivotField', sheet: d.t.sheet, pivot: d.t.pivot, field: d.field, group })
  }
  /**
   * Excel's Group Selection: the Grouping dialog over a field of dates or numbers with one item selected, otherwise the
   * items selected of one field gathered into a group by hand.
   */
  async function group(): Promise<boolean> {
    const got = await picked()
    if (!got) return false
    const items = got.parts.filter((x) => x.kind === 'item' && cellField(x) !== undefined && x.item !== undefined)
    const field = cellField(items[0])
    if (field === undefined || items.some((x) => x.field !== field)) return false
    const unique = [...new Set(items.map((x) => x.item!))]
    if (fieldKind(got.t, field) !== 'text' && unique.length === 1) return openGroup(got.t, field)
    const out = await send({ op: 'groupPivotField', sheet: got.t.sheet, pivot: got.t.pivot, field, group: { by: 'items', items: unique } })
    return !!out
  }
  /** Group Field: Grouping over the active cell's field. */
  async function groupField(): Promise<boolean> {
    const got = await active()
    const field = cellField(got?.part)
    return got && field !== undefined ? openGroup(got.t, field) : false
  }
  /** Excel's Ungroup: the groups selected of a field grouped by hand, or the active cell's field of dates or numbers. */
  async function ungroup(): Promise<EditOutcome | null> {
    const got = await picked()
    const part = got?.parts.find((x) => cellField(x) !== undefined)
    if (!got || !part) return null
    const field = ungroupedField(got.t, cellField(part)!)
    if (field === null) return null
    const base = { op: 'ungroupPivotField' as const, sheet: got.t.sheet, pivot: got.t.pivot, field }
    if (!got.t.def.cache.fields[field]?.group?.discrete) return send(base)
    // the groups selected of a field grouped by hand, each item of theirs a group of its own again
    const items = [...new Set(got.parts.filter((x) => x.field === field && x.item !== undefined).map((x) => x.item!))]
    return send({ ...base, ...(items.length ? { items } : {}) })
  }

  // ── Calculated Field and Item ──
  const calcNames = (t: PivotTarget) => t.def.cache.fields.flatMap((f) => (f.formula !== undefined ? [f.name] : []))
  /** Excel's first free name of a new one: Field1, Field2… or Formula1… in the workbook's language. */
  function freeName(taken: string[], name: (n: number) => string): string {
    const used = new Set(taken.map((x) => x.toLowerCase()))
    let n = 1
    while (used.has(name(n).toLowerCase())) n++
    return name(n)
  }
  function openCalcField(t: PivotTarget): void {
    const name = freeName(t.def.cache.fields.map((f) => f.name), m().pivotCalcFieldName)
    open({ kind: 'calcField', t, draft: { name, formula: '= 0', pick: 0, item: 0 } })
  }
  function openCalcItem(t: PivotTarget, field: number): boolean {
    if (!t.cells.labels[field]) return false
    const name = freeName(t.cells.labels[field] ?? [], m().pivotCalcItemName)
    open({ kind: 'calcItem', t, field, draft: { name, formula: '= 0', pick: field, item: 0 } })
    return true
  }
  /** Add or Modify, and OK with a name: the calculated field or item set (a new calculated field placed on Values too, as Excel's OK does). */
  async function applyCalc(closing: boolean): Promise<EditOutcome | null> {
    const d = state.value
    if ((d?.kind !== 'calcField' && d?.kind !== 'calcItem') || !d.draft.name.trim()) {
      if (closing) close()
      return null
    }
    const name = d.draft.name.trim()
    const formula = await canonical(d.draft.formula)
    const t = d.t
    if (closing) close()
    if (d.kind === 'calcItem') return kept(await send({ op: 'setPivotCalculatedItem', sheet: t.sheet, pivot: t.pivot, field: d.field, name, formula }))
    // Excel's OK puts a new calculated field on Values after the others with Sum, in the same step (xlcalc's `values`); Modify leaves them
    const known = calcNames(t).some((x) => x.toLowerCase() === name.toLowerCase())
    return kept(await send({ op: 'setPivotCalculatedField', sheet: t.sheet, pivot: t.pivot, name, formula, ...(known ? {} : { values: true }) }))
  }
  /** After Add or Modify the dialog stays open over the PivotTable read again, its names and formulas listed anew. */
  async function kept(out: EditOutcome | null): Promise<EditOutcome | null> {
    const d = state.value
    if (out?.ok && (d?.kind === 'calcField' || d?.kind === 'calcItem')) {
      const again = await target(d.t.sheet, d.t.pivot)
      if (again && state.value === d) d.t = markRaw(again.t)
    }
    return out
  }
  /**
   * A formula typed in the workbook's locale as xlcalc's operations take it (`=Amount*0,5` in ru-RU as `=Amount*0.5`), read
   * by xlcalc's entry parsing; as typed in an en-US workbook, or where the parsing refuses it (a field's name in quotes), for
   * xlcalc to answer.
   */
  async function canonical(text: string): Promise<string> {
    const typed = text.trim().startsWith('=') ? text.trim() : `=${text.trim()}`
    if (s.locale.value.toLowerCase().startsWith('en')) return typed
    const got = (await loadEntry()).parseEntry(typed, { locale: deps.locale() })
    return got.ok && got.input?.formula !== undefined ? `=${got.input.formula}` : typed
  }

  /** Delete: the calculated field or item named in the box. */
  async function deleteCalc(): Promise<EditOutcome | null> {
    const d = state.value
    if ((d?.kind !== 'calcField' && d?.kind !== 'calcItem') || !d.draft.name.trim()) return null
    const name = d.draft.name.trim()
    close()
    if (d.kind === 'calcItem') return send({ op: 'setPivotCalculatedItem', sheet: d.t.sheet, pivot: d.t.pivot, field: d.field, name, formula: null })
    return send({ op: 'setPivotCalculatedField', sheet: d.t.sheet, pivot: d.t.pivot, name, formula: null })
  }

  // ── the commands a host's ribbon calls ──
  async function command(c: 'options' | 'calculatedField' | 'calculatedItem' | 'groupField'): Promise<boolean> {
    if (c === 'groupField') return groupField()
    const got = await active()
    if (!got) return false
    if (c === 'options') openOptions(got.t)
    else if (c === 'calculatedField') openCalcField(got.t)
    else {
      const f = cellField(got.part)
      return f !== undefined && got.part?.axis !== 'pages' ? openCalcItem(got.t, f) : false
    }
    return true
  }
  async function design(change: PivotDesignChange): Promise<EditOutcome | null> {
    const got = await active()
    return got ? edit(got.t, designPatch(got.t, change)) : null
  }

  return {
    can,
    locale: deps.locale,
    state,
    close,
    views,
    here,
    target,
    picked,
    edit,
    send,
    openValue,
    pickFunction,
    pickBaseField,
    applyValue,
    setFunction,
    setShowAs,
    openNumberFormat,
    applyNumberFormat,
    openField,
    applyField,
    openOptions,
    optionsValid,
    applyOptions,
    openGroup,
    applyGroup,
    group,
    groupField,
    ungroup,
    calcNames,
    openCalcField,
    openCalcItem,
    applyCalc,
    deleteCalc,
    command,
    design,
  }
}

export type PivotDialogs = ReturnType<typeof usePivotDialogs>
