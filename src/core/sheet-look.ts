// The sheet's look beside its cells: its tab's color, and Excel's View > Show (gridlines, headings), kept through the
// source's `show`, or in the sheet's storage for a source without it or a read-only sheet.
import { errorText } from '../failure'
import type { SheetInfo, SheetShow } from '../contract/source'
import type { SheetState } from './state'

/** Whether a `#RRGGBB` color is dark, by its perceived brightness: a tab filled with it takes light text. */
export function darkColour(hex: string): boolean {
  const n = Number.parseInt(hex.slice(1, 7), 16)
  if (Number.isNaN(n)) return false
  const r = (n >> 16) & 255
  const g = (n >> 8) & 255
  const b = n & 255
  return (r * 299 + g * 587 + b * 114) / 1000 < 128
}

export type ShowSetting = keyof SheetShow

export function useSheetLook(s: SheetState, notify: (text: string) => void) {
  const memoryName = (sheet: string) => `show:${sheet}`

  /** The sheets as read, with the settings the sheet keeps itself for a source that cannot; a source that keeps them wins on an editable sheet. */
  function apply(list: SheetInfo[]): SheetInfo[] {
    if (s.editable.value && s.sourceMethod('show')) return list
    return list.map((x) => {
      const own = s.memory.recall<SheetShow>(memoryName(x.name))
      return own ? { ...x, ...own } : x
    })
  }

  /** Puts a setting on a sheet in the list as displayed, by name, since the list may have been read again meanwhile. */
  function put(sheet: string, what: ShowSetting, on: boolean): void {
    s.sheets.value = s.sheets.value.map((x) => (x.name === sheet ? { ...x, [what]: on } : x))
  }

  /** Excel's View > Show on the displayed sheet, toggled without `on`: drawn at once, then kept; whether it shows afterwards. */
  function show(what: ShowSetting, on?: boolean): boolean {
    const info = s.sheet.value
    if (!info) return false
    const was = info[what] !== false
    const next = on ?? !was
    if (next === was) return next
    const sheet = info.name
    put(sheet, what, next)
    const call = s.editable.value ? s.sourceMethod('show') : null
    if (!call) {
      const own = { ...s.memory.recall<SheetShow>(memoryName(sheet)), [what]: next }
      s.memory.remember(memoryName(sheet), own)
      return next
    }
    // The source keeps it now: the browser's copy, stored while the sheet was read-only, goes
    const kept = s.memory.recall<SheetShow>(memoryName(sheet))
    if (kept && what in kept) {
      const { [what]: _, ...rest } = kept
      s.memory.remember(memoryName(sheet), Object.keys(rest).length ? rest : null)
    }
    const outdated = s.outdated()
    // Not an undo step, as in Excel: a refusal puts the setting back
    call(sheet, { [what]: next }).then(
      (out) => out.ok || refused(out.message),
      (e: unknown) => refused(errorText(e)),
    )
    function refused(text: string): void {
      if (outdated()) return
      put(sheet, what, was)
      notify(text || s.messages.value.editFailed)
    }
    return next
  }

  return { apply, show }
}

export type SheetLook = ReturnType<typeof useSheetLook>
