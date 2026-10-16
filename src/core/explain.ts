// Excel's Trace Precedents as a panel: the source's tree of the active cell's precedents as text, asked again as the cell
// moves or its value changes (after a pause, a late answer dropped), its references to go to.
import { computed, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { cellAddress, quoteSheetIfNeeded } from 'xlcalc/a1'
import type { GridReference } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'
import type { Formula } from './formula'
import type { SheetState } from './state'

/** The pause after the active cell moves before the source is asked, so keys held down do not ask for every cell they pass. */
export const EXPLAIN_DELAY_MS = 150
/** What sets a line's formula off from its node: xlcalc's `explain` writes `B4 = 30  ← SUM(B2:B3)`. */
const ARROW = '  ← '

/** A piece of a line of the tree: a reference with the sheet it is on, or text, `from` when it is after the node's arrow (its formula, where the value came from). */
export interface ExplainPart {
  text: string
  ref?: GridReference & { sheet: string }
  from?: boolean
}
export interface ExplainLine {
  depth: number
  parts: ExplainPart[]
}

/**
 * The source's text as the tree's lines, as xlcalc's `explain` writes it: a node a line, two spaces of indent a level, the
 * node's reference first (`'Q 4'!A2 = 60  ← Data!A4*2`). References to cells, columns or rows of one sheet become parts,
 * read as stored formula text, which reads addresses alike in every locale; one without a sheet is on its line's node's
 * sheet, else on `sheet`. Anything else (a table, a reference across sheets, another workbook's) stays text.
 */
export function explainLines(text: string, sheet: string): ExplainLine[] {
  const out: ExplainLine[] = []
  for (const raw of text.split(/\r?\n/)) {
    const body = raw.trimStart()
    if (!body.trim()) continue
    const arrow = body.indexOf(ARROW)
    const parts: ExplainPart[] = []
    // Text from `a` to `b`, cut where the arrow begins.
    const plain = (a: number, b: number) => {
      const cut = arrow > a && arrow < b ? arrow : b
      if (cut > a) parts.push({ text: body.slice(a, cut), ...(arrow >= 0 && a >= arrow ? { from: true } : {}) })
      if (cut < b) parts.push({ text: body.slice(cut, b), from: true })
    }
    const refs = scanReferences(body).filter((r): r is GridReference => r.kind !== 'table' && r.sheetLast === undefined)
    const own = refs[0]?.at === 0 ? refs[0].sheet : undefined
    let at = 0
    for (const r of refs) {
      plain(at, r.at)
      parts.push({ text: r.text, ref: { ...r, sheet: r.sheet ?? own ?? sheet } })
      at = r.end
    }
    plain(at, body.length)
    out.push({ depth: Math.floor((raw.length - body.length) / 2), parts })
  }
  return out
}

export function useExplain(s: SheetState, f: Formula) {
  const available = computed(() => typeof s.source().explain === 'function')
  const open = ref(false)
  /** The cell the lines are of (`'Q 4'!B6`), the lines, an answer awaited, the last one failed. */
  const place = ref('')
  const lines = shallowRef<ExplainLine[]>([])
  const busy = ref(false)
  const failed = ref(false)
  let timer: ReturnType<typeof setTimeout> | null = null
  /** The request in force: an older answer, or one asked before the panel closed, is dropped (the contract's `explain` takes no signal). */
  let seq = 0

  /** Ask for the active cell after `delay` (at once without one); a request still out is dropped when its answer comes. */
  function schedule(delay = EXPLAIN_DELAY_MS): void {
    if (timer) clearTimeout(timer)
    timer = null
    const mine = ++seq
    if (!open.value) return
    busy.value = true
    if (delay > 0) timer = setTimeout(() => void ask(mine), delay)
    else void ask(mine)
  }
  async function ask(mine: number): Promise<void> {
    timer = null
    const a = s.selected.value
    const call = s.source().explain
    if (!a || !call || !s.sheet.value) {
      place.value = ''
      lines.value = []
      busy.value = false
      return
    }
    const sheet = s.sheetName.value
    const cell = cellAddress(a.row, a.col)
    let got: ExplainLine[] = []
    let bad = false
    try {
      got = explainLines(await call.call(s.source(), sheet, cell), sheet)
    } catch {
      bad = true
    }
    if (mine !== seq) return
    place.value = `${quoteSheetIfNeeded(sheet)}!${cell}`
    lines.value = got
    failed.value = bad
    busy.value = false
  }

  // The active cell, and what it shows: a move, or a value an edit changed, asks again while the panel is open.
  const subject = computed(() => {
    const a = s.selected.value
    return a && s.sheet.value ? [s.sheetName.value, a.row, a.col, f.selectedFormula.value, f.selectedValue.value].join('\u0000') : ''
  })
  watch(subject, () => {
    if (open.value) schedule()
  })

  /** Open the panel at the active cell; false without the source's `explain` or a cell. */
  function show(): boolean {
    if (!available.value || !s.sheet.value || !s.selected.value) return false
    open.value = true
    schedule(0)
    return true
  }
  function close(): void {
    if (!open.value) return
    open.value = false
    schedule()
    busy.value = false
    s.focusGrid()
  }
  /** The book changed elsewhere: the tree is asked again. */
  function refresh(): void {
    if (open.value) schedule()
  }
  onScopeDispose(() => {
    if (timer) clearTimeout(timer)
  })

  return { available, open, place, lines, busy, failed, show, close, refresh }
}

export type Explain = ReturnType<typeof useExplain>
