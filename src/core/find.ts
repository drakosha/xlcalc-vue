// Excel's Find and Replace on the sheet: the bar's query and options go to the source's `find`, Find Next continues from
// the active cell in the chosen order, a selection of several cells limits the search, and Replace and Replace All go to
// the source's `replace` as an edit does.
import { computed, nextTick, onScopeDispose, reactive, ref, watch } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { CellPosition } from 'xlcalc/contract'
import { isAbort } from '../abort'
import { cellOrder } from '../find-match'
import type { EditOutcome, FindRequest, ReplaceOutcome } from '../contract/source'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** A pause after typing before the source is queried. */
const FIND_DELAY_MS = 250

/** The options of Excel's Find dialog. */
export interface FindOptions {
  matchCase: boolean
  wholeCell: boolean
  lookIn: 'formulas' | 'values'
  byColumns: boolean
}

/** A change sent to the source as an edit is sent (`Editing.run`): its outcome applied to the windows, a rejection reported. */
export type FindRunner = (call: () => Promise<EditOutcome>) => Promise<EditOutcome>

export function useFind(s: SheetState, sel: Selection, notify: (text: string) => void) {
  const available = computed(() => typeof s.source().find === 'function')
  const canReplace = computed(() => available.value && s.editable.value && typeof s.source().replace === 'function')
  const isOpen = ref(false)
  /** The Replace row is shown: Ctrl+H, or its toggle on the bar. */
  const replacing = ref(false)
  const optionsOpen = ref(false)
  const text = ref('')
  const by = ref('')
  const options = reactive<FindOptions>({ matchCase: false, wholeCell: false, lookIn: 'formulas', byColumns: false })
  /** Excel's Replace searches in formulas whatever Find was set to. */
  const lookIn = computed(() => (replacing.value ? 'formulas' : options.lookIn))
  /** The range searched, as Excel searches a selection of more than one cell: taken when the bar opens or such a range is selected. */
  const scope = ref<string | null>(null)
  const hits = ref<CellPosition[]>([])
  const total = ref(0)
  /** The cell the matches were requested from (Find After), null when they start from the range's first cell. */
  const from = ref<string | null>(null)
  const busy = ref(false)
  /** A replacement is pending: the Replace buttons wait. */
  const writing = ref(false)
  /** What Replace All did, shown in the status until the search changes. */
  const done = ref<string | null>(null)
  const input = ref<HTMLInputElement | null>(null)
  let timer: ReturnType<typeof setTimeout> | null = null
  let seq = 0
  let ctrl: AbortController | null = null
  /** A search requested by the query or the options goes to its first match after the active cell. */
  let jump = false

  /** Matches by `row:col` (sheet rows): those in the viewport are highlighted, the current one pulses. */
  const hitSet = computed(() => new Set(hits.value.map((h) => `${h.row}:${h.col}`)))
  /** Excel's order: row by row, or column by column; negative when `a` comes first. */
  const order = (a: CellPosition, b: CellPosition) => cellOrder(options.byColumns)(a, b)
  /** The match at the active cell, -1 when the active cell is not a match. */
  const index = computed(() => {
    const a = s.selected.value
    return a ? hits.value.findIndex((h) => h.row === a.row && h.col === a.col) : -1
  })
  const current = computed(() => {
    const h = hits.value[index.value]
    return h ? `${h.row}:${h.col}` : ''
  })

  function open(replace: boolean): void {
    if (!available.value) return
    isOpen.value = true
    replacing.value = replace && canReplace.value
    scope.value = sel.multi.value ? sel.selectionRef.value : null
    // the last query comes back with the bar: its matches are counted, and the viewport stays
    if (text.value) schedule(false)
    void nextTick(() => {
      input.value?.focus()
      input.value?.select()
    })
  }
  /** Ctrl+F: the bar without the Replace row. */
  const openFind = (): void => open(false)
  /** Ctrl+H: the bar with the Replace row, or without it where the source does not replace. */
  const openReplace = (): void => open(true)
  function closeFind(): void {
    isOpen.value = false
    reset()
    s.focusGrid()
  }
  function reset(): void {
    seq++
    ctrl?.abort()
    if (timer) clearTimeout(timer)
    hits.value = []
    total.value = 0
    from.value = null
    busy.value = false
    done.value = null
  }
  // The sheet is unmounted: no search after the pause, none pending.
  onScopeDispose(reset)

  /** The bar's current request, from the cell `after` when given; null for an empty query. */
  function request(after?: string): FindRequest | null {
    if (!text.value || !s.sheet.value) return null
    return {
      ...(after ? { after } : {}),
      sheet: s.sheetName.value,
      query: text.value,
      ...(scope.value ? { range: scope.value } : {}),
      ...(options.matchCase ? { matchCase: true } : {}),
      ...(options.wholeCell ? { wholeCell: true } : {}),
      lookIn: lookIn.value,
      ...(options.byColumns ? { byColumns: true } : {}),
      ...(s.visible.value ? { filter: s.visible.value.spec } : {}),
    }
  }

  /**
   * Queries the source now, from the range's first cell or from `after`; a newer search or a reset discards the response. A
   * new query goes to its first match after the active cell: among the returned matches when they are all there, otherwise
   * requested from that cell.
   */
  async function runSearch(after?: string): Promise<void> {
    if (timer) clearTimeout(timer)
    timer = null
    const req = request(after)
    if (!req || !isOpen.value || !available.value) return reset()
    const mine = ++seq
    const leap = jump
    jump = false
    ctrl?.abort()
    const c = (ctrl = new AbortController())
    busy.value = true
    try {
      const got = await s.source().find!(req, c.signal)
      if (mine !== seq) return
      hits.value = got.hits.map((h) => ({ row: h.row, col: h.col }))
      total.value = got.total
      from.value = after ?? null
      const a = s.selected.value
      if (after) show(0)
      // more matches than returned: the first one after the active cell is requested from it
      else if (leap && a && got.total > got.hits.length) return void runSearch(cellAddress(a.row, a.col))
      else if (leap) step(1)
    } catch (e) {
      if (mine === seq && !isAbort(e)) notify(s.messages.value.findFailed)
    } finally {
      if (mine === seq) busy.value = false
    }
  }
  /** Queries the source after the pause; `leap`: go to the first match after the active cell once it responds. */
  function schedule(leap: boolean): void {
    jump ||= leap
    seq++
    if (timer) clearTimeout(timer)
    if (!text.value || !isOpen.value) return reset()
    busy.value = true
    timer = setTimeout(() => void runSearch(), FIND_DELAY_MS)
  }
  // A new query or option goes to its first match; another sheet, filter or range, or the Replace row's toggle (Replace searches
  // in formulas), only recounts, and the viewport and the active cell stay.
  watch([text, () => ({ ...options })], () => {
    done.value = null
    schedule(true)
  })
  watch([() => s.sheetName.value, () => s.visible.value?.key ?? '', scope, lookIn], () => schedule(false))
  // A range selected while the bar is open limits the next search; selecting single matched cells leaves it as it is.
  watch(sel.selectionRef, (range) => {
    if (isOpen.value && sel.multi.value && range !== scope.value) scope.value = range
  })

  function show(i: number): void {
    const h = hits.value[i]
    if (!h) return
    s.selected.value = { row: h.row, col: h.col }
    s.extent.value = null
    void nextTick(() => sel.revealCell(h.row, h.col))
  }
  /**
   * Find Next and Find Previous, as Excel moves, wrapping around at the end: from a match to its neighbor in the source's
   * order (Find's, by blocks on a filtered sheet), from another cell to the first match after it in the chosen order. Beyond
   * the matches the source returned (a few hundred at a time), Find Next requests the next ones from the active cell, as
   * Excel's Find After does.
   */
  function step(dir: 1 | -1): void {
    const list = hits.value
    if (!list.length) return
    done.value = null
    const a = s.selected.value
    if (!a) return show(dir > 0 ? 0 : list.length - 1)
    const i = index.value
    const some = total.value > list.length
    if (dir > 0 && some && (i < 0 || i === list.length - 1)) return void runSearch(cellAddress(a.row, a.col))
    if (i >= 0) return show((i + dir + list.length) % list.length)
    if (dir > 0) {
      const i = list.findIndex((h) => order(h, a) > 0)
      return show(i >= 0 ? i : 0)
    }
    for (let i = list.length - 1; i >= 0; i--) if (order(list[i]!, a) < 0) return show(i)
    show(list.length - 1)
  }

  /** Excel's Replace: the active cell when it is a match, then Find Next; elsewhere only Find Next. */
  async function replaceOne(run: FindRunner): Promise<void> {
    const req = request()
    const call = s.sourceMethod('replace')
    const a = s.selected.value
    if (!req || !call || !canReplace.value || writing.value) return
    if (index.value < 0 || !a) return step(1)
    writing.value = true
    try {
      const out = await run(() => call({ ...req, range: cellAddress(a.row, a.col) }, by.value))
      if (!out.ok) return
      await runSearch()
      step(1)
    } finally {
      writing.value = false
    }
  }
  /** Excel's Replace All over what the bar finds, as a single undo step; the bar reports how many cells changed and counts again. */
  async function replaceAll(run: FindRunner): Promise<void> {
    const req = request()
    const call = s.sourceMethod('replace')
    if (!req || !call || !canReplace.value || writing.value) return
    writing.value = true
    try {
      const out = await run(() => call(req, by.value))
      if (!out.ok) return
      const { replaced, refused } = out as ReplaceOutcome
      // the cells Excel would not accept are left unchanged and the others written: the message line says why
      if (refused?.length) notify(s.messages.value.replaceRefused(refused.length, refused[0]!.message))
      await runSearch()
      done.value = s.messages.value.replaceDone(replaced ?? 0)
    } finally {
      writing.value = false
    }
  }

  function onFindKey(e: KeyboardEvent): void {
    if (e.key === 'Enter') {
      e.preventDefault()
      step(e.shiftKey ? -1 : 1)
    } else if (e.key === 'Escape') {
      e.preventDefault()
      closeFind()
    }
  }
  const status = computed(() => {
    if (done.value) return done.value
    if (!text.value) return ''
    if (busy.value) return '…'
    const m = s.messages.value
    if (!hits.value.length) return m.findNone
    // a position among the matches is known only when they were requested from the range's first cell
    return index.value >= 0 && !from.value ? m.findStatus(index.value + 1, total.value, hits.value.length) : m.findCount(total.value, hits.value.length)
  })

  return {
    available,
    canReplace,
    isOpen,
    replacing,
    optionsOpen,
    text,
    by,
    options,
    lookIn,
    scope,
    hits,
    total,
    index,
    busy,
    writing,
    input,
    hitSet,
    current,
    openFind,
    openReplace,
    closeFind,
    step,
    replaceOne,
    replaceAll,
    onFindKey,
    status,
  }
}

export type Find = ReturnType<typeof useFind>
