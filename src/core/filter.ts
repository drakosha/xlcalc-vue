// The display filter and sort: the data do not change, and the source lists the rows to display, in order. Its funnels are the
// buttons on the header row of the sheet's AutoFilter, over the AutoFilter's rows; without one, on the column letters.
import { computed, nextTick, onScopeDispose, ref, watch } from 'vue'
import { colLetters, parseRange, sameName, MAX_COLS } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import { isAbort } from '../abort'
import { filterKey } from '../filter-key'
import { clamp } from '../rect'
import type { SheetFilter, SheetOp } from '../contract/source'
import type { SheetState } from './state'
import type { Windows } from './windows'

type ColumnFilter = NonNullable<SheetFilter['cols']>[number]
type Sort = NonNullable<SheetFilter['sort']>

export function useFilter(s: SheetState, w: Windows, notify: (text: string) => void) {
  /** The conditions and the sort as set; those on columns without a funnel (outside the AutoFilter) are kept but not in effect. */
  const own = ref<Record<number, ColumnFilter>>({})
  const ownSort = ref<Sort | null>(null)
  /** The funnel needs both: the rows to display and a column's values. */
  const available = computed(() => typeof s.source().visibleRows === 'function' && typeof s.source().columnValues === 'function')
  /** The sheet's AutoFilter (`SheetInfo.autoFilter`), null without one. */
  const auto = computed<CellRect | null>(() => {
    const r = s.sheet.value?.autoFilter ? parseRange(s.sheet.value.autoFilter.range) : null
    return r?.kind === 'cells' ? r : null
  })
  /** The source can put an AutoFilter on the sheet (Data > Filter): a sheet without one then has no funnels, as in Excel. */
  const settable = computed(() => s.editable.value && !!s.sourceMethod('setAutoFilter'))
  /** The funnels are on the column letters: no AutoFilter on the sheet, and the source cannot set one. */
  const letters = computed(() => available.value && !auto.value && !settable.value)
  /** A column has a funnel: within the AutoFilter, or any column with the funnels on the letters. */
  const canFunnel = (c: number): boolean => (auto.value ? available.value && c >= auto.value.left && c <= auto.value.right : letters.value)
  /** The conditions and the sort in effect; setting them sets them as given. */
  const cols = computed({
    get: () => {
      const out: Record<number, ColumnFilter> = {}
      for (const [c, f] of Object.entries(own.value)) if (canFunnel(Number(c))) out[Number(c)] = f
      return out
    },
    set: (next: Record<number, ColumnFilter>) => void (own.value = next),
  })
  const sort = computed({ get: () => (ownSort.value && canFunnel(ownSort.value.col) ? ownSort.value : null), set: (next: Sort | null) => void (ownSort.value = next) })
  /** Header rows the filter leaves alone: those down to the AutoFilter's header row; otherwise the workbook's frozen rows, or the first row, as Excel does. */
  const head = computed(() => {
    if (auto.value) return auto.value.top
    const frozen = s.sheet.value?.freeze?.rows ?? 0
    return clamp(frozen > 0 ? frozen : 1, 0, Math.max(0, s.dataRows.value - 1))
  })
  const active = computed(() => Object.keys(cols.value).length > 0 || sort.value !== null)
  /** The rows the conditions apply to: below the header, down to the AutoFilter's last row. */
  const rowsSpec = (): SheetFilter => (auto.value ? { head: head.value, last: auto.value.bottom } : { head: head.value })
  const spec = computed<SheetFilter | null>(() => {
    if (!active.value) return null
    const out = rowsSpec()
    if (Object.keys(cols.value).length) out.cols = cols.value
    if (sort.value) out.sort = sort.value
    return out
  })
  const key = computed(() => (spec.value ? filterKey(spec.value) : ''))

  // ── the list of displayed rows: requested when the conditions change, and windows wait for it ──
  let seq = 0
  let ctrl: AbortController | null = null
  /** Puts a list of displayed rows into effect (null: none); windows start over in the new rows. */
  function apply(next: SheetState['visible']['value']): void {
    s.visible.value = next
    s.filterWait.value = false
    w.reset()
    w.ensure()
  }
  async function fetchVisible(): Promise<void> {
    const mine = ++seq
    ctrl?.abort()
    const want = spec.value
    const wantKey = key.value
    remember()
    if (!want || !available.value) {
      if (s.visible.value || s.filterWait.value) apply(null)
      return
    }
    if (s.visible.value?.key === wantKey) return
    s.filterWait.value = true
    ctrl = new AbortController()
    try {
      const got = await s.source().visibleRows!(s.sheetName.value, want, ctrl.signal)
      if (mine !== seq) return
      apply(got.rows ? { key: wantKey, spec: want, rows: Int32Array.from(got.rows), index: displayIndex(got.rows) } : null)
    } catch (e) {
      if (mine !== seq || isAbort(e)) return
      notify(s.messages.value.filterFailed)
      clear()
    }
  }
  watch([() => s.sheetName.value, key], () => void fetchVisible())
  /** Each sheet row's display index (0: not displayed), with a row listed twice at its last position: a typed array, since a million rows take 4 MB. */
  function displayIndex(rows: number[]): Int32Array {
    const index = new Int32Array(rows.reduce((m, r) => Math.max(m, r), 0) + 1)
    rows.forEach((r, i) => (index[r] = i + 1))
    return index
  }

  /** Conditions are stored per sheet (under the host's storage key). */
  const memoryName = () => `filter:${s.sheetName.value}`
  function remember(): void {
    if (s.sheetName.value) s.memory.remember(memoryName(), active.value ? { cols: cols.value, sort: sort.value } : null)
  }
  /** Moves the conditions and the sort to the columns `moved` gives, dropping those it gives none. */
  function remap(moved: (c: number) => number | null): void {
    const next: Record<number, ColumnFilter> = {}
    for (const [c, f] of Object.entries(own.value)) {
      const to = moved(Number(c))
      if (to !== null) next[to] = f
    }
    own.value = next
    const by = ownSort.value && moved(ownSort.value.col)
    ownSort.value = ownSort.value && by ? { ...ownSort.value, col: by } : null
  }
  const sameRect = (a: CellRect, b: CellRect) => a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right
  /** The AutoFilter's change the watcher followed, with the conditions before it, for a structure command whose hook comes after. */
  let followed: { before: CellRect; now: CellRect; own: Record<number, ColumnFilter>; sort: Sort | null } | null = null
  /** The AutoFilter's change the sheet's own structure command leads to, whose conditions `moveColumns` has moved already. */
  let moved: { before: CellRect; now: CellRect | null } | null = null
  // An AutoFilter removed (Filter off, its header row deleted, a table over it) takes its conditions along, as in Excel.
  // One moved otherwise (an undo, a redo, another viewer) takes them along when it moved whole; those outside it go.
  watch([() => s.sheetName.value, auto], ([sheet, now], [was, before]) => {
    const mine = moved
    moved = null
    followed = null
    if (sheet !== was || !before) return
    if (!now) return clear()
    if (sameRect(before, now) || (mine && mine.now && sameRect(mine.before, before) && sameRect(mine.now, now))) return
    followed = { before, now, own: own.value, sort: ownSort.value }
    const shift = now.right - now.left === before.right - before.left ? now.left - before.left : 0
    remap((c) => (c + shift >= now.left && c + shift <= now.right ? c + shift : null))
  })
  /** The AutoFilter's range after columns are inserted or deleted, as Excel adjusts it; null when every column of it is deleted. */
  function rangeAfter(r: CellRect, insert: boolean, at: number, n: number): CellRect | null {
    const end = (c: number, left: boolean) => (c < at ? c : insert ? c + n : c >= at + n ? c - n : left ? at : at - 1)
    const out = { ...r, left: end(r.left, true), right: end(r.right, false) }
    return out.right < out.left ? null : out
  }
  /**
   * Columns inserted or deleted on the sheet by its own structure command: the conditions and the sort move with their
   * columns, and those of deleted columns go, as Excel moves an AutoFilter's criteria.
   */
  function moveColumns(op: SheetOp): void {
    if ((op.op !== 'insertColumns' && op.op !== 'deleteColumns') || !sameName(op.sheet, s.sheetName.value)) return
    const at = parseRange(`${op.column}1`)?.left ?? MAX_COLS + 1
    const n = op.count ?? 1
    const insert = op.op === 'insertColumns'
    // The sheets may have been read again before this hook: the watcher then followed this very change, which is undone here.
    const f = followed
    followed = null
    const done = f && auto.value && sameRect(auto.value, f.now) && rangeAfter(f.before, insert, at, n)
    if (f && done && sameRect(done, f.now)) {
      own.value = f.own
      ownSort.value = f.sort
    } else if (auto.value) moved = { before: auto.value, now: rangeAfter(auto.value, insert, at, n) }
    remap((c) => (c < at ? c : insert ? c + n : c < at + n ? null : c - n))
  }
  /** The sheet's stored conditions, before its windows load; the row list arrives through the watcher. */
  function restore(): void {
    close()
    const saved = s.memory.recall<{ cols?: Record<number, ColumnFilter>; sort?: Sort | null }>(memoryName())
    s.visible.value = null
    own.value = saved?.cols ?? {}
    ownSort.value = saved?.sort ?? null
    s.filterWait.value = active.value
  }
  /** Reads the list again (the source reloaded): the rows may have changed. */
  function refresh(): Promise<void> {
    if (s.visible.value) s.visible.value = { ...s.visible.value, key: '' }
    return fetchVisible()
  }
  /**
   * Excel's Reapply (Ctrl+Alt+L): applies the conditions and the sort again to the current data, so rows that no longer
   * match after an edit are hidden and rows that now match appear; the scroll stays, and so does the active cell unless its
   * row was hidden, in which case the row now at its position takes it. False when no filter is in effect or the list could not be read.
   */
  async function reapply(): Promise<boolean> {
    if (!active.value || !available.value) return false
    const a = s.selected.value
    const at = a ? (s.indexOf(a.row) ?? null) : null
    const sheet = s.sheetName.value
    await refresh()
    if (!s.visible.value || s.visible.value.key !== key.value || s.sheetName.value !== sheet) return false
    const cur = s.selected.value
    const shown = s.visible.value.rows.length
    if (cur && at !== null && shown && s.indexOf(cur.row) === undefined) {
      s.selected.value = { row: s.rowOf(Math.min(at, shown)), col: cur.col }
      s.extent.value = null
    }
    return true
  }
  /** Clears every condition at once: the rows return to the sheet's order. */
  function clear(): void {
    own.value = {}
    ownSort.value = null
    seq++
    ctrl?.abort()
    apply(null)
    remember()
  }
  s.hooks.unfilter = (row: number) => {
    clear()
    notify(s.messages.value.unfiltered(row))
  }

  /** The status line below the grid: visible rows out of all (of the AutoFilter's rows below its header), the columns with conditions, the sort. */
  const status = computed(() => {
    if (!active.value) return ''
    const fmt = new Intl.NumberFormat(s.locale.value)
    const last = auto.value ? Math.min(auto.value.bottom, s.dataRows.value) : s.dataRows.value
    const rows = s.visible.value?.rows
    // the rows below the AutoFilter come last in the list, in order
    let below = 0
    if (rows) while (below < rows.length && rows[rows.length - 1 - below]! > last) below++
    const shown = Math.max(0, (rows ? rows.length - below : last) - head.value)
    const names = Object.keys(cols.value).map(Number).sort((a, b) => a - b).map(colLetters)
    const sorted = sort.value ? { col: colLetters(sort.value.col), dir: sort.value.dir } : null
    return s.messages.value.filterStatus(fmt.format(shown), fmt.format(Math.max(0, last - head.value)), names, sorted)
  })

  // ── a column's funnel: its values with checkboxes, a substring, the sort ──
  const open = ref<number | null>(null)
  const pos = ref({ left: 0, top: 0 })
  const values = ref<{ text: string; count: number }[]>([])
  const more = ref(0)
  const busy = ref(false)
  /** The draft: checked values (null: all) and a substring; Apply or Enter puts it into effect. */
  const draftValues = ref<Set<string> | null>(null)
  const draftText = ref('')
  const draftSearch = ref('')
  let valuesSeq = 0
  async function openFunnel(col: number, anchor?: DOMRect): Promise<void> {
    if (open.value === col) return close()
    pos.value = { left: Math.max(4, Math.min(anchor?.left ?? 0, (typeof window !== 'undefined' ? window.innerWidth : 800) - 292)), top: (anchor?.bottom ?? 0) + 2 }
    open.value = col
    const cur = cols.value[col]
    draftValues.value = cur?.values ? new Set(cur.values) : null
    draftText.value = cur?.text ?? ''
    draftSearch.value = ''
    values.value = []
    more.value = 0
    busy.value = true
    const mine = ++valuesSeq
    try {
      const got = await s.source().columnValues!(s.sheetName.value, col, spec.value ?? rowsSpec())
      if (mine !== valuesSeq) return
      values.value = got.values
      more.value = got.more
    } catch {
      if (mine === valuesSeq) notify(s.messages.value.valuesFailed)
    } finally {
      if (mine === valuesSeq) busy.value = false
    }
  }
  function close(): void {
    if (open.value === null) return
    open.value = null
    valuesSeq++
    void nextTick(s.focusGrid)
  }
  /** The funnel's title: the header cell's text and the column letters, or the letters alone. */
  const title = computed(() => {
    const c = open.value
    if (c === null) return ''
    const name = head.value > 0 ? w.textOfRow(head.value, c).trim() : ''
    return name ? s.messages.value.filterNamed(name, colLetters(c)) : s.messages.value.filterColumn(colLetters(c))
  })
  /** The list narrowed by its own search box (a convenience, not a condition, as in Excel). */
  const shownValues = computed(() => {
    const q = draftSearch.value.trim().toLocaleLowerCase()
    return q ? values.value.filter((v) => v.text.toLocaleLowerCase().includes(q)) : values.value
  })
  const isChecked = (text: string) => draftValues.value === null || draftValues.value.has(text)
  const allChecked = computed(() => draftValues.value === null || shownValues.value.every((v) => draftValues.value!.has(v.text)))
  function toggleValue(text: string): void {
    const next = new Set(draftValues.value ?? values.value.map((v) => v.text))
    if (next.has(text)) next.delete(text)
    else next.add(text)
    draftValues.value = next
  }
  function toggleAll(): void {
    if (allChecked.value) {
      const next = new Set(draftValues.value ?? values.value.map((v) => v.text))
      for (const v of shownValues.value) next.delete(v.text)
      draftValues.value = next
    } else {
      const next = new Set(draftValues.value ?? [])
      for (const v of shownValues.value) next.add(v.text)
      draftValues.value = next.size === values.value.length && values.value.every((v) => next.has(v.text)) ? null : next
    }
  }
  /** Puts the draft into effect: a condition on the column, or none when everything is checked and there is no substring. */
  function applyDraft(): void {
    const c = open.value
    if (c === null) return
    const cond: ColumnFilter = {}
    const all = draftValues.value === null || (values.value.length > 0 && values.value.every((v) => draftValues.value!.has(v.text)))
    if (!all && draftValues.value) cond.values = [...draftValues.value]
    if (draftText.value.trim()) cond.text = draftText.value.trim()
    const next = { ...cols.value }
    if (cond.values || cond.text) next[c] = cond
    else delete next[c]
    own.value = next
    close()
  }
  function clearColumn(c: number): void {
    const next = { ...cols.value }
    delete next[c]
    own.value = next
    if (sort.value?.col === c) ownSort.value = null
    close()
  }
  function sortBy(c: number, dir: 'asc' | 'desc'): void {
    ownSort.value = sort.value?.col === c && sort.value.dir === dir ? null : { col: c, dir }
    close()
  }
  /** A ribbon's Sort A to Z or Z to A on a column: that sort, whatever the current one. */
  function setSort(c: number, dir: 'asc' | 'desc'): void {
    ownSort.value = { col: c, dir }
    close()
  }
  function onFunnelKey(e: KeyboardEvent): void {
    if (e.key === 'Enter') {
      e.preventDefault()
      applyDraft()
    } else if (e.key === 'Escape') {
      e.preventDefault()
      close()
    }
  }

  // The sheet is unmounted: the pending list is aborted, and pending column values are discarded.
  onScopeDispose(() => {
    seq++
    valuesSeq++
    ctrl?.abort()
  })

  return {
    cols,
    sort,
    available,
    auto,
    letters,
    canFunnel,
    moveColumns,
    head,
    active,
    spec,
    status,
    restore,
    refresh,
    reapply,
    clear,
    open,
    pos,
    values,
    more,
    busy,
    draftText,
    draftSearch,
    title,
    shownValues,
    isChecked,
    allChecked,
    openFunnel,
    close,
    toggleValue,
    toggleAll,
    applyDraft,
    clearColumn,
    sortBy,
    setSort,
    onFunnelKey,
  }
}

export type Filter = ReturnType<typeof useFilter>
