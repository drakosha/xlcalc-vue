// The display filter and sort, the funnel on a column's letter: the data do not change, the source lists the rows to show in order.
import { computed, nextTick, onScopeDispose, ref, watch } from 'vue'
import { colLetters } from 'xlcalc/a1'
import { isAbort } from '../abort'
import { filterKey } from '../filter-key'
import { clamp } from '../rect'
import type { SheetFilter } from '../source'
import type { SheetState } from './state'
import type { Windows } from './windows'

type ColumnFilter = NonNullable<SheetFilter['cols']>[number]
type Sort = NonNullable<SheetFilter['sort']>

export function useFilter(s: SheetState, w: Windows, notify: (text: string) => void) {
  const cols = ref<Record<number, ColumnFilter>>({})
  const sort = ref<Sort | null>(null)
  /** The funnel needs both: the rows to show and a column's values. */
  const available = computed(() => typeof s.source().visibleRows === 'function' && typeof s.source().columnValues === 'function')
  /** Header rows the filter leaves alone: the book's frozen rows, else the first row, as Excel does. */
  const head = computed(() => {
    const frozen = s.sheet.value?.freeze?.rows ?? 0
    return clamp(frozen > 0 ? frozen : 1, 0, Math.max(0, s.dataRows.value - 1))
  })
  const active = computed(() => Object.keys(cols.value).length > 0 || sort.value !== null)
  const spec = computed<SheetFilter | null>(() => {
    if (!active.value) return null
    const out: SheetFilter = { head: head.value }
    if (Object.keys(cols.value).length) out.cols = cols.value
    if (sort.value) out.sort = sort.value
    return out
  })
  const key = computed(() => (spec.value ? filterKey(spec.value) : ''))

  // ── the list of shown rows: asked when the conditions change, windows wait for it ──
  let seq = 0
  let ctrl: AbortController | null = null
  /** Put a list of shown rows in force (null: none); windows start over in the new rows. */
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
  /** Each sheet row's display index (0: not shown), a row listed twice at its last place: a typed array, a million rows are 4 MB. */
  function displayIndex(rows: number[]): Int32Array {
    const index = new Int32Array(rows.reduce((m, r) => Math.max(m, r), 0) + 1)
    rows.forEach((r, i) => (index[r] = i + 1))
    return index
  }

  /** Conditions are remembered by sheet (under the host's storage key). */
  const memoryName = () => `filter:${s.sheetName.value}`
  function remember(): void {
    if (s.sheetName.value) s.memory.remember(memoryName(), active.value ? { cols: cols.value, sort: sort.value } : null)
  }
  /** The sheet's conditions from memory, before its windows load; the list of rows comes by the watcher. */
  function restore(): void {
    close()
    const saved = s.memory.recall<{ cols?: Record<number, ColumnFilter>; sort?: Sort | null }>(memoryName())
    s.visible.value = null
    cols.value = saved?.cols ?? {}
    sort.value = saved?.sort ?? null
    s.filterWait.value = !!(available.value && (Object.keys(cols.value).length || sort.value))
  }
  /** Read the list again (the source reloaded): the rows may have changed. */
  function refresh(): Promise<void> {
    if (s.visible.value) s.visible.value = { ...s.visible.value, key: '' }
    return fetchVisible()
  }
  /**
   * Excel's Reapply (Control+Alt+L): the conditions and the sort again over the data as they are now, so rows an edit took
   * out of them go and rows it brought in come; the scroll stays, and the active cell too unless its row went, when the row
   * now at its place takes it. False when no filter is in force or the list could not be read.
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
  /** Drop every condition at once: the rows go back to the sheet's order. */
  function clear(): void {
    cols.value = {}
    sort.value = null
    seq++
    ctrl?.abort()
    apply(null)
    remember()
  }
  s.hooks.unfilter = (row: number) => {
    clear()
    notify(s.messages.value.unfiltered(row))
  }

  /** The line under the grid: shown rows of all, the columns with conditions, the sort. */
  const status = computed(() => {
    if (!active.value) return ''
    const fmt = new Intl.NumberFormat(s.locale.value)
    const shown = Math.max(0, (s.visible.value?.rows.length ?? s.dataRows.value) - head.value)
    const names = Object.keys(cols.value).map(Number).sort((a, b) => a - b).map(colLetters)
    const sorted = sort.value ? { col: colLetters(sort.value.col), dir: sort.value.dir } : null
    return s.messages.value.filterStatus(fmt.format(shown), fmt.format(Math.max(0, s.dataRows.value - head.value)), names, sorted)
  })

  // ── the funnel of a column: its values with checks, a substring, the sort ──
  const open = ref<number | null>(null)
  const pos = ref({ left: 0, top: 0 })
  const values = ref<{ text: string; count: number }[]>([])
  const more = ref(0)
  const busy = ref(false)
  /** The draft: checked values (null: all) and a substring; Apply or Enter puts it in force. */
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
      const got = await s.source().columnValues!(s.sheetName.value, col, spec.value ?? { head: head.value })
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
  /** The funnel's title: the header cell's text and the letters, or the letters alone. */
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
  /** Put the draft in force: a condition on the column, or none when all is checked and there is no substring. */
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
    cols.value = next
    close()
  }
  function clearColumn(c: number): void {
    const next = { ...cols.value }
    delete next[c]
    cols.value = next
    if (sort.value?.col === c) sort.value = null
    close()
  }
  function sortBy(c: number, dir: 'asc' | 'desc'): void {
    sort.value = sort.value?.col === c && sort.value.dir === dir ? null : { col: c, dir }
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

  // The sheet's end: the list out is aborted, a column's values out dropped.
  onScopeDispose(() => {
    seq++
    valuesSeq++
    ctrl?.abort()
  })

  return {
    cols,
    sort,
    available,
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
    onFunnelKey,
  }
}

export type Filter = ReturnType<typeof useFilter>
