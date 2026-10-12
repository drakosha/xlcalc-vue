// Search on the sheet, as the the app's: the source finds (a substring in any case), the grid steps through the hits.
import { computed, nextTick, ref, watch } from 'vue'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** A pause after typing before the source is asked. */
const FIND_DELAY_MS = 250

export function useFind(s: SheetState, sel: Selection, notify: (text: string) => void) {
  const available = computed(() => typeof s.props.source.find === 'function')
  const isOpen = ref(false)
  const text = ref('')
  const hits = ref<{ row: number; col: number }[]>([])
  const total = ref(0)
  const index = ref(-1)
  const busy = ref(false)
  const input = ref<HTMLInputElement | null>(null)
  let timer: ReturnType<typeof setTimeout> | null = null
  let seq = 0
  let ctrl: AbortController | null = null
  /** Hits by `row:col` (sheet rows): those in the frame are tinted, the current one pulses. */
  const hitSet = computed(() => new Set(hits.value.map((h) => `${h.row}:${h.col}`)))
  const current = computed(() => {
    const h = hits.value[index.value]
    return h ? `${h.row}:${h.col}` : ''
  })

  function openFind(): void {
    if (!available.value) return
    isOpen.value = true
    void nextTick(() => {
      input.value?.focus()
      input.value?.select()
    })
  }
  function closeFind(): void {
    isOpen.value = false
    reset()
    s.scroller.value?.focus({ preventScroll: true })
  }
  function reset(): void {
    seq++
    ctrl?.abort()
    if (timer) clearTimeout(timer)
    hits.value = []
    total.value = 0
    index.value = -1
    busy.value = false
  }
  // A new query jumps to its first hit; another sheet or filter only recounts the hits, the frame stays.
  watch([text, () => s.sheetName.value, () => s.visible.value?.key ?? ''], ([q], [prevQ]) => {
    const query = q.trim()
    if (!query || !isOpen.value || !available.value) return reset()
    const jump = q !== prevQ
    const mine = ++seq
    ctrl?.abort()
    if (timer) clearTimeout(timer)
    busy.value = true
    timer = setTimeout(async () => {
      const c = (ctrl = new AbortController())
      try {
        const got = await s.props.source.find!({ sheet: s.sheetName.value, query, filter: s.visible.value?.spec }, c.signal)
        if (mine !== seq) return
        hits.value = got.hits.map((h) => ({ row: h.row, col: h.col }))
        total.value = got.total
        index.value = got.hits.length ? 0 : -1
        if (got.hits.length && jump) show(0)
      } catch (e) {
        if (mine === seq && !(e instanceof DOMException && e.name === 'AbortError')) notify(s.messages.value.findFailed)
      } finally {
        if (mine === seq) busy.value = false
      }
    }, FIND_DELAY_MS)
  })
  function show(i: number): void {
    const h = hits.value[i]
    if (!h) return
    index.value = i
    s.selected.value = { row: h.row, col: h.col }
    s.extent.value = null
    void nextTick(() => sel.revealCell(h.row, h.col))
  }
  function step(dir: 1 | -1): void {
    const n = hits.value.length
    if (n) show((index.value + dir + n) % n)
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
    if (!text.value.trim()) return ''
    if (busy.value) return '…'
    if (!hits.value.length) return s.messages.value.findNone
    return s.messages.value.findStatus(index.value + 1, total.value, hits.value.length)
  })

  return { available, isOpen, text, hits, total, index, busy, input, hitSet, current, openFind, closeFind, step, onFindKey, status }
}

export type Find = ReturnType<typeof useFind>
