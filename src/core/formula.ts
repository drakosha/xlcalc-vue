// The active cell's formula in pieces with clickable references, and the formula's ranges highlighted on the grid as in Excel.
import { computed, type Ref } from 'vue'
import { colNumber } from './a1'
import { formulaTokens, refKey, type FormulaRef, type FormulaToken } from './formula-refs'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** Colours of a formula's ranges, as Excel's: each range its own, equal references one. */
const REF_COLORS = ['#2563eb', '#dc2626', '#7c3aed', '#16a34a', '#ea580c', '#0891b2', '#db2777']
/** How many formulas of a selection are read and how many ranges painted: past that it is noise and slow. */
const PRECEDENT_CELLS_MAX = 500
const PRECEDENT_RANGES_MAX = 24

export function useFormula(s: SheetState, w: Windows, sel: Selection, draft: Ref<string | null>) {
  const formulaAt = (row: number, col: number): string | undefined => w.formulaOfRow(row, col)
  const selectedFormula = computed(() => (s.selected.value ? (formulaAt(s.selected.value.row, s.selected.value.col) ?? '') : ''))
  const selectedValue = computed(() => (s.selected.value ? w.textOfRow(s.selected.value.row, s.selected.value.col) : ''))
  /** The formula being typed, when the draft is one, takes over the highlights. */
  const shownFormula = computed(() => {
    const d = draft.value
    if (d !== null) return d.startsWith('=') ? d.slice(1) : ''
    return selectedFormula.value
  })

  /** Source ranges: the active cell's formula first (its colours in the bar), then the other formulas of the selection. */
  const refColors = computed(() => {
    const out = new Map<string, { ref: FormulaRef; color: string }>()
    const consider = (formula: string) => {
      for (const t of formulaTokens(formula)) {
        if (!t.ref) continue
        const key = refKey(t.ref)
        if (out.has(key)) continue
        if (out.size >= PRECEDENT_RANGES_MAX) return
        out.set(key, { ref: t.ref, color: REF_COLORS[out.size % REF_COLORS.length]! })
      }
    }
    if (shownFormula.value) consider(shownFormula.value)
    const r = sel.selRect.value
    if (r && sel.multi.value && draft.value === null) {
      let cells = 0
      for (const kind of ['body', 'top', 'left'] as const) {
        const slot = w.slots[kind].value
        if (!slot.win || slot.view !== w.viewKey()) continue
        for (const [addr, f] of Object.entries(slot.win.formulas)) {
          const m = /^([A-Z]+)(\d+)$/.exec(addr)
          if (!m) continue
          const row = Number(m[2])
          if (row < r.r1 || row > r.r2) continue
          const col = colNumber(m[1]!)
          if (col < r.c1 || col > r.c2 || sel.isActive(row, col)) continue
          if (++cells > PRECEDENT_CELLS_MAX) return out
          consider(f)
        }
      }
    }
    return out
  })
  /** Pieces of the active cell's formula; a reference carries its range and colour. */
  const formulaParts = computed<(FormulaToken & { color?: string })[]>(() => {
    if (!selectedFormula.value) return []
    return formulaTokens(selectedFormula.value).map((t) => (t.ref ? { ...t, color: refColors.value.get(refKey(t.ref))?.color } : t))
  })
  /** The sheet a reference leads to: its index, or -1 when the book has no such sheet. */
  function refSheet(ref: { sheet: string | null }): number {
    if (ref.sheet === null) return s.active.value
    return sheetIndexByName(ref.sheet)
  }
  function sheetIndexByName(name: string): number {
    const low = name.toLowerCase()
    return s.sheets.value.findIndex((x) => x.name.toLowerCase() === low)
  }
  /** Highlighted ranges on the active sheet. */
  const highlights = computed(() => [...refColors.value.values()].filter(({ ref }) => refSheet(ref) === s.active.value))
  /** A cell's highlight: a frame on the range's edges only (Excel draws no inner lines) and a light tint. */
  function highlightAt(row: number, col: number): { edges: string[]; tints: string[] } | null {
    const hl = highlights.value
    if (!hl.length) return null
    const edges: string[] = []
    const tints: string[] = []
    for (const { ref, color } of hl) {
      const r2 = Math.min(ref.r2, s.rowCount.value)
      const c2 = Math.min(ref.c2, s.colCount.value)
      if (row < ref.r1 || row > r2 || col < ref.c1 || col > c2) continue
      if (row === ref.r1) edges.push(`inset 0 2px 0 0 ${color}`)
      if (row === r2) edges.push(`inset 0 -2px 0 0 ${color}`)
      if (col === ref.c1) edges.push(`inset 2px 0 0 0 ${color}`)
      if (col === c2) edges.push(`inset -2px 0 0 0 ${color}`)
      tints.push(`linear-gradient(${color}1f, ${color}1f)`)
    }
    return tints.length ? { edges, tints } : null
  }

  return { formulaAt, selectedFormula, selectedValue, refColors, formulaParts, refSheet, sheetIndexByName, highlights, highlightAt }
}

export type Formula = ReturnType<typeof useFormula>
