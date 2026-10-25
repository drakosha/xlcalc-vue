// Excel's Show Formulas (Ctrl+`, Formulas > Show Formulas): a sheet's view setting under which cells show their formulas
// and constants without their number formats, and columns are drawn twice as wide; the workbook does not change. Kept
// per sheet in the sheet's storage, as the frozen panes choice is, since the source has no place for it.
import { shallowRef } from 'vue'
import { loadEntry, type Entry } from '../entry'
import type { Formula } from './formula'
import type { SheetState } from './state'
import type { CellData } from './windows'

export function useFormulaView(s: SheetState, f: Formula) {
  const memoryName = () => `formulas:${s.sheetName.value}`
  /** xlcalc's typed-entry module, which writes a number as the formula bar does; null until it loads. */
  const entry = shallowRef<Entry | null>(null)
  function need(): void {
    if (!entry.value) loadEntry().then((m) => (entry.value = m), () => {})
  }
  /** The displayed sheet's setting, read when the sheet is displayed. */
  function read(): void {
    s.formulaView.value = s.memory.recall<boolean>(memoryName()) === true
    if (s.formulaView.value) need()
  }
  /** Turns Show Formulas on or off on the displayed sheet (toggles it without `on`); whether it is on afterwards. */
  function show(on = !s.formulaView.value): boolean {
    if (!s.sheet.value) return false
    s.formulaView.value = on
    s.memory.remember(memoryName(), on || null)
    if (on) need()
    return on
  }
  /** Ctrl+` (Control on a Mac too, as in Excel for Mac; the key's position, with any layout); true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (!e.ctrlKey || e.altKey || e.metaKey || e.shiftKey || (e.code !== 'Backquote' && e.key !== '`')) return false
    e.preventDefault()
    show()
    return true
  }

  /**
   * What a cell shows under Show Formulas: its formula with `=`; a number as the formula bar writes it in General (in the
   * workbook's locale, so a date is its serial number); text as stored; anything else as displayed. Until xlcalc/entry
   * has loaded, a number shows its displayed text.
   */
  function textOf(data: CellData): string {
    if (data.formula !== undefined) return `=${data.formula}`
    const value = data.value
    if (typeof value === 'string') return value
    const lib = entry.value
    const opts = f.scanOptions()
    if (typeof value !== 'number' || !lib || !opts) return data.text
    try {
      return lib.entryText(value, { locale: opts.locale })
    } catch {
      return data.text
    }
  }

  return { read, show, key, textOf }
}

export type FormulaView = ReturnType<typeof useFormulaView>
