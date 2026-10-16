// Excel's Find and Replace on the sheet: the bar's query and options go to the source's `find`, Find Next goes on from
// the active cell in the order asked, a selection of several cells is where it looks, and Replace and Replace All go to
// the source's `replace` as an edit goes.
import { computed, nextTick, reactive, ref, watch } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { CellPosition } from 'xlcalc/contract'
import type { EditOutcome, FindRequest, ReplaceOutcome } from '../source'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** A pause after typing before the source is asked. */
const FIND_DELAY_MS = 250

/** The options of Excel's Find dialog. */
export interface FindOptions {
  matchCase: boolean
  wholeCell: boolean
  lookIn: 'formulas' | 'values'
  byColumns: boolean
}

/** A change sent to the source as an edit is sent (`Editing.run`): its outcome applied to the windows, a refusal told. */
export type FindRunner = (call: () => Promise<EditOutcome>) => Promise<EditOutcome>

export function useFind(s: SheetState, sel: Selection, notify: (text: string) => void) {
  const available = computed(() => typeof s.source().find === 'function')
  const canReplace = computed(() => available.value && s.editable.value && typeof s.source().replace === 'function')
  const isOpen = ref(false)
  /** The replace row shows: Ctrl+H, or its toggle on the bar. */
  const replacing = ref(false)
  const optionsOpen = ref(false)
  const text = ref('')
  const by = ref('')
  const options = reactive<FindOptions>({ matchCase: false, wholeCell: false, lookIn: 'formulas', byColumns: false })
  /** Excel's Replace looks in formulas whatever Find was set to. */
  const lookIn = computed(() => (replacing.value ? 'formulas' : options.lookIn))
  /** The range searched, as Excel searches a selection of more than one cell: taken when the bar opens or such a range is selected. */
  const scope = ref<string | null>(null)
  const hits = ref<CellPosition[]>([])
  const total = ref(0)
  /** The cell the hits were asked from (Find After), null when they come from the range's first cell. */
  const from = ref<string | null>(null)
  const busy = ref(false)
  /** A replacement is out: the replace buttons wait. */
  const writing = ref(false)
  /** What Replace All did, in the status until the search changes. */
  const done = ref<string | null>(null)
  const input = ref<HTMLInputElement | null>(null)
  let timer: ReturnType<typeof setTimeout> | null = null
  let seq = 0
  let ctrl: AbortController | null = null
  /** A search asked for by the query or the options goes to its first hit after the active cell. */
  let jump = false

  /** Hits by `row:col` (sheet rows): those in the frame are tinted, the active one pulses. */
  const hitSet = computed(() => new Set(hits.value.map((h) => `${h.row}:${h.col}`)))
  /** Excel's order: row by row, or column by column; negative when `a` comes first. */
  const order = (a: CellPosition, b: CellPosition) => (options.byColumns ? a.col - b.col || a.row - b.row : a.row - b.row || a.col - b.col)
  /** The hit at the active cell, -1 when the active cell is none. */
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
    // the last query comes back with the bar: its hits are counted, the frame stays
    if (text.value) schedule(false)
    void nextTick(() => {
      input.value?.focus()
      input.value?.select()
    })
  }
  /** Ctrl+F: the bar without the replace row. */
  const openFind = (): void => open(false)
  /** Ctrl+H: the bar with the replace row, or without it where the source does not replace. */
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

  /** The request of the bar as it is now, from the cell `after` when given; null for an empty query. */
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
   * Ask the source now, from the range's first cell or from `after`; a newer search or a reset drops the answer. A new query
   * goes to its first hit after the active cell: among the hits given when they are all, else asked from that cell.
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
      // more hits than given: the first after the active cell is asked from it
      else if (leap && a && got.total > got.hits.length) return void runSearch(cellAddress(a.row, a.col))
      else if (leap) step(1)
    } catch (e) {
      if (mine === seq && !(e instanceof DOMException && e.name === 'AbortError')) notify(s.messages.value.findFailed)
    } finally {
      if (mine === seq) busy.value = false
    }
  }
  /** Ask the source after the pause; `leap`: go to the first hit after the active cell once it answers. */
  function schedule(leap: boolean): void {
    jump ||= leap
    seq++
    if (timer) clearTimeout(timer)
    if (!text.value || !isOpen.value) return reset()
    busy.value = true
    timer = setTimeout(() => void runSearch(), FIND_DELAY_MS)
  }
  // A new query or option goes to its first hit; another sheet, filter or range, or the replace row's toggle (Replace looks
  // in formulas), only recounts, the frame and the active cell stay.
  watch([text, () => ({ ...options })], () => {
    done.value = null
    schedule(true)
  })
  watch([() => s.sheetName.value, () => s.visible.value?.key ?? '', scope, lookIn], () => schedule(false))
  // A range selected while the bar is open is where it looks next; the single cells of the hits leave it as it is.
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
   * Find Next and Find Previous, as Excel goes, round past the end: from a hit to its neighbour in the source's order (Find's,
   * by blocks on a filtered sheet), from another cell to the first hit past it in the order asked. Past the hits the source
   * gave (it gives some hundreds of them), Find Next asks for the next ones from the active cell, as Excel's Find After.
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

  /** Excel's Replace: the active cell when it is a hit, then Find Next; elsewhere only Find Next. */
  async function replaceOne(run: FindRunner): Promise<void> {
    const req = request()
    const call = s.source().replace
    const a = s.selected.value
    if (!req || !call || !canReplace.value || writing.value) return
    if (index.value < 0 || !a) return step(1)
    writing.value = true
    try {
      const out = await run(() => call.call(s.source(), { ...req, range: cellAddress(a.row, a.col) }, by.value))
      if (!out.ok) return
      await runSearch()
      step(1)
    } finally {
      writing.value = false
    }
  }
  /** Excel's Replace All over what the bar finds, one step of undo; the bar says how many cells it changed and counts again. */
  async function replaceAll(run: FindRunner): Promise<void> {
    const req = request()
    const call = s.source().replace
    if (!req || !call || !canReplace.value || writing.value) return
    writing.value = true
    try {
      const out = await run(() => call.call(s.source(), req, by.value))
      if (!out.ok) return
      const { replaced, refused } = out as ReplaceOutcome
      // the cells Excel would not take are left as they were, the others written: the line says why
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
    // a place among the hits is known only when they were asked from the range's first cell
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
