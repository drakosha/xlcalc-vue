// The active cell's formula in pieces with clickable references, and the formula's ranges highlighted on the grid as in Excel.
import { computed, shallowRef, watch, type Ref } from 'vue'
import { nameKey, parseCell, sameName } from 'xlcalc/a1'
import type { GridReference, ScanOptions } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'
import { LOCALE_MODULES } from '../locales'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** A piece of a formula as the bar shows it; a reference carries its range and colour. */
export interface FormulaPart {
  text: string
  ref?: GridReference
  color?: string
}

/** Colours of a formula's ranges, as Excel's: each range its own, equal references one. */
const REF_COLORS = ['#2563eb', '#dc2626', '#7c3aed', '#16a34a', '#ea580c', '#0891b2', '#db2777']
/** How many formulas of a selection are read and how many ranges painted: past that it is noise and slow. */
const PRECEDENT_CELLS_MAX = 500
const PRECEDENT_RANGES_MAX = 24

/** A reference's key: equal references (`A1`, `$A$1`; a sheet in any case) get one colour. */
const refKey = (r: GridReference): string => `${r.sheet === undefined ? '' : nameKey(r.sheet)}!${r.top}:${r.left}:${r.bottom}:${r.right}`

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

  /** How formula text is read: in the book's locale once its module is in, as stored text where xlcalc has none; null while it loads. */
  const scan = shallowRef<ScanOptions | null>(null)
  watch(
    s.locale,
    (tag) => {
      const load = LOCALE_MODULES[tag]
      scan.value = load ? null : {}
      if (!load) return
      // A module that fails to load leaves the text read as stored: addresses read alike in every locale.
      const take = (opts: ScanOptions) => {
        if (s.locale.value === tag) scan.value = opts
      }
      void load().then(
        (locale) => take({ locale }),
        () => take({}),
      )
    },
    { immediate: true },
  )
  /** References the grid shows: to cells, columns or rows of one sheet; a table's and one across sheets stay text. */
  function gridRefs(formula: string): GridReference[] {
    const opts = scan.value
    if (!opts) return []
    return scanReferences(formula, opts).filter((r): r is GridReference => r.kind !== 'table' && r.sheetLast === undefined)
  }

  /** Source ranges: the active cell's formula first (its colours in the bar), then the other formulas of the selection. */
  const refColors = computed(() => {
    const out = new Map<string, { ref: GridReference; color: string }>()
    const consider = (formula: string) => {
      for (const ref of gridRefs(formula)) {
        const key = refKey(ref)
        if (out.has(key)) continue
        if (out.size >= PRECEDENT_RANGES_MAX) return
        out.set(key, { ref, color: REF_COLORS[out.size % REF_COLORS.length]! })
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
          const at = parseCell(addr)
          if (!at || at.row < r.top || at.row > r.bottom || at.col < r.left || at.col > r.right || sel.isActive(at.row, at.col)) continue
          if (++cells > PRECEDENT_CELLS_MAX) return out
          consider(f)
        }
      }
    }
    return out
  })
  /** Pieces of the active cell's formula: the text between references, and each reference with its range and colour. */
  const formulaParts = computed<FormulaPart[]>(() => {
    const formula = selectedFormula.value
    if (!formula) return []
    const out: FormulaPart[] = []
    let at = 0
    for (const ref of gridRefs(formula)) {
      if (ref.at > at) out.push({ text: formula.slice(at, ref.at) })
      out.push({ text: ref.text, ref, color: refColors.value.get(refKey(ref))?.color })
      at = ref.end
    }
    if (at < formula.length) out.push({ text: formula.slice(at) })
    return out
  })
  /** The sheet a reference leads to: its index, or -1 when the book has no such sheet. */
  function refSheet(ref: { sheet?: string }): number {
    if (ref.sheet === undefined) return s.active.value
    return sheetIndexByName(ref.sheet)
  }
  function sheetIndexByName(name: string): number {
    return s.sheets.value.findIndex((x) => sameName(x.name, name))
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
      const bottom = Math.min(ref.bottom, s.rowCount.value)
      const right = Math.min(ref.right, s.colCount.value)
      if (row < ref.top || row > bottom || col < ref.left || col > right) continue
      if (row === ref.top) edges.push(`inset 0 2px 0 0 ${color}`)
      if (row === bottom) edges.push(`inset 0 -2px 0 0 ${color}`)
      if (col === ref.left) edges.push(`inset 2px 0 0 0 ${color}`)
      if (col === right) edges.push(`inset -2px 0 0 0 ${color}`)
      tints.push(`linear-gradient(${color}1f, ${color}1f)`)
    }
    return tints.length ? { edges, tints } : null
  }

  return { formulaAt, selectedFormula, selectedValue, refColors, formulaParts, refSheet, sheetIndexByName, highlights, highlightAt }
}

export type Formula = ReturnType<typeof useFormula>
