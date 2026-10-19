// The formula the formula bar displays (the draft's while one is typed) in pieces with colored references, and the formula's ranges highlighted on the grid, as in Excel.
import { computed, type Ref } from 'vue'
import { cellAddress, nameKey, sameName } from 'xlcalc/a1'
import type { CellRect, GridReference, ScanOptions } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'
import { isGridRef } from '../grid-ref'
import { useLocaleModule } from '../locale-ref'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** A piece of a formula as the formula bar displays it; a reference carries its range and color. */
export interface FormulaPart {
  text: string
  ref?: GridReference
  color?: string
}

/** The colors of a formula's ranges, as in Excel: each range its own, and equal references the same one. */
const REF_COLORS = ['#2563eb', '#dc2626', '#7c3aed', '#16a34a', '#ea580c', '#0891b2', '#db2777']
/** How many formulas of a selection are read and how many ranges painted: beyond that it is noise and slow. */
const PRECEDENT_CELLS_MAX = 500
const PRECEDENT_RANGES_MAX = 24
/** The number of formula texts whose references are cached; beyond that the cache starts over. */
const SCAN_CACHE_MAX = 4096

/** A reference's key: equal references (`A1`, `$A$1`; a sheet name in any letter case) get one color. */
const refKey = (r: GridReference): string => `${r.sheet === undefined ? '' : nameKey(r.sheet)}!${r.top}:${r.left}:${r.bottom}:${r.right}`

/** Whether two color maps hold the same references with the same colors in the same order. */
function sameColors(a: Map<string, { color: string }>, b: Map<string, { color: string }>): boolean {
  if (a.size !== b.size) return false
  const it = a.entries()
  for (const [key, { color }] of b) {
    const [k, x] = it.next().value!
    if (k !== key || x.color !== color) return false
  }
  return true
}

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

  /** How formula text is parsed: in the workbook's locale once its module has loaded, as stored text where xlcalc has none or it failed to load (addresses parse the same in every locale); null while it loads. */
  const bookLocale = useLocaleModule(() => s.locale.value)
  const scan = computed<ScanOptions | null>(() => (bookLocale.value === null ? null : bookLocale.value ? { locale: bookLocale.value } : {}))
  /** How the workbook's formula text is parsed, for `xlcalc/text`; null while its locale loads. */
  const scanOptions = (): ScanOptions | null => scan.value
  /** References of formulas parsed under the current scan options, by text: a drag reads the same cells over and over. */
  let scanned = new Map<string, GridReference[]>()
  let scannedWith: ScanOptions | null = null
  /** References the grid displays: to cells, columns or rows of a single sheet; a table reference and a 3D reference stay text. */
  function gridRefs(formula: string): GridReference[] {
    const opts = scan.value
    if (!opts) return []
    if (scannedWith !== opts || scanned.size >= SCAN_CACHE_MAX) {
      scanned = new Map()
      scannedWith = opts
    }
    let refs = scanned.get(formula)
    if (!refs) {
      refs = scanReferences(formula, opts).filter(isGridRef)
      scanned.set(formula, refs)
    }
    return refs
  }

  /** The formulas of the selection's cells except the active one, by window (body, frozen rows, frozen columns) and row by row; each cell once, up to the limit. */
  function selectionFormulas(r: CellRect, take: (formula: string) => void): void {
    // Cells change in place: a computed that reads them follows the windows' revision.
    void w.revision.value
    let cells = 0
    for (const kind of ['body', 'top', 'left'] as const) {
      const slot = w.slots[kind].value
      const win = slot.win
      if (!win || slot.view !== w.viewKey()) continue
      // Without a filter display rows are the sheet's rows, so rows outside the selection are not walked.
      const d1 = s.visible.value ? win.row : Math.max(win.row, r.top)
      const d2 = s.visible.value ? win.row + slot.rows - 1 : Math.min(win.row + slot.rows - 1, r.bottom)
      const c1 = Math.max(win.col, r.left)
      const c2 = Math.min(win.col + slot.cols - 1, r.right)
      for (let d = d1; d <= d2; d++) {
        const row = s.rowOf(d)
        if (row < r.top || row > r.bottom) continue
        for (let col = c1; col <= c2; col++) {
          const f = win.formulas[cellAddress(row, col)]
          // A cell held by two windows belongs to the first one.
          if (f === undefined || sel.isActive(row, col) || w.slotAt(d, col) !== slot) continue
          if (++cells > PRECEDENT_CELLS_MAX) return
          take(f)
        }
      }
    }
  }
  /** Source ranges: the active cell's formula first (its colors in the formula bar), then the selection's other formulas; the same map as before while they are the same. */
  const refColors = computed((prev?: Map<string, { ref: GridReference; color: string }>) => {
    const out = new Map<string, { ref: GridReference; color: string }>()
    const consider = (formula: string) => {
      if (out.size >= PRECEDENT_RANGES_MAX) return
      for (const ref of gridRefs(formula)) {
        const key = refKey(ref)
        if (out.has(key)) continue
        if (out.size >= PRECEDENT_RANGES_MAX) return
        out.set(key, { ref, color: REF_COLORS[out.size % REF_COLORS.length]! })
      }
    }
    if (shownFormula.value) consider(shownFormula.value)
    const r = sel.selRect.value
    if (r && sel.multi.value && draft.value === null) selectionFormulas(r, consider)
    return prev && sameColors(prev, out) ? prev : out
  })
  /** Pieces of the displayed formula (the draft's while one is typed, otherwise the active cell's): the text between references, and each reference with its range and the color of its border on the grid. */
  const formulaParts = computed<FormulaPart[]>(() => {
    const formula = shownFormula.value
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
  /** The sheet a reference points to: its index, or -1 when the workbook has no such sheet. */
  function refSheet(ref: { sheet?: string }): number {
    if (ref.sheet === undefined) return s.active.value
    return sheetIndexByName(ref.sheet)
  }
  function sheetIndexByName(name: string): number {
    return s.sheets.value.findIndex((x) => sameName(x.name, name))
  }
  /** Highlighted ranges on the active sheet, clipped to the grid, with their border sides and tint prepared for each cell. */
  const highlights = computed(() =>
    [...refColors.value.values()]
      .filter(({ ref }) => refSheet(ref) === s.active.value)
      .map(({ ref, color }) => ({
        ref,
        color,
        bottom: Math.min(ref.bottom, s.rowCount.value),
        right: Math.min(ref.right, s.colCount.value),
        top: `inset 0 2px 0 0 ${color}`,
        under: `inset 0 -2px 0 0 ${color}`,
        start: `inset 2px 0 0 0 ${color}`,
        end: `inset -2px 0 0 0 ${color}`,
        tint: `linear-gradient(${color}1f, ${color}1f)`,
      })),
  )
  /** A cell's highlight: a border on the range's edges only (Excel draws no inner lines) and a light tint; `hl` the ranges to search (a row's). */
  function highlightAt(row: number, col: number, hl: readonly Highlight[] = highlights.value): { edges: string[]; tints: string[] } | null {
    let out: { edges: string[]; tints: string[] } | null = null
    for (const h of hl) {
      if (row < h.ref.top || row > h.bottom || col < h.ref.left || col > h.right) continue
      out ??= { edges: [], tints: [] }
      if (row === h.ref.top) out.edges.push(h.top)
      if (row === h.bottom) out.edges.push(h.under)
      if (col === h.ref.left) out.edges.push(h.start)
      if (col === h.right) out.edges.push(h.end)
      out.tints.push(h.tint)
    }
    return out
  }

  return { formulaAt, selectedFormula, selectedValue, shownFormula, scanOptions, formulaParts, refSheet, sheetIndexByName, highlights, highlightAt }
}

export type Formula = ReturnType<typeof useFormula>
/** A highlighted range with its border sides and tint. */
export type Highlight = Formula['highlights']['value'][number]
