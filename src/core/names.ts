// The workbook's defined names, read from the source once per source and again after a change (Define Name, the Name
// Manager, undo, redo, a reload): which are visible on a sheet, the range each refers to, and replacing the list.
import { computed, onScopeDispose, shallowRef } from 'vue'
import { sameName } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import { failedOutcome } from '../failure'
import type { CellUpdate, EditOutcome } from '../contract/source'
import type { SheetName } from '../contract/workbook'
import { askSource } from './ask'
import { parsePlace } from './place'
import type { SheetState } from './state'

/** The range a name refers to: a plain reference to a range of one sheet. */
export interface NameRange {
  sheet: string
  rect: CellRect
}

export interface NamesDeps {
  /** Reads the sheets and windows again, and the number of reloads started, to tell whether the source sent `reload` during a call. */
  reload: () => Promise<void>
  reloads: () => number
  /** Cells whose values a change of names changed. */
  apply: (changes: CellUpdate[]) => void
}

/** A name's identity in the list: its scope and its name, case-insensitively, as Excel tells names apart. */
export const nameKey = (n: SheetName): string => `${n.scope?.toLocaleUpperCase() ?? ''}!${n.name.toLocaleUpperCase()}`

/** `Sheet!Name` or `'My sheet'!Name`: a name of a sheet's scope typed with its sheet. */
const SCOPED = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/

export function useNames(s: SheetState, deps: NamesDeps) {
  /** The workbook's names as the source last returned them; empty without the source's `names`. */
  const list = shallowRef<readonly SheetName[]>([])
  /** The source the list was read from: a new source starts empty. */
  let readFrom: unknown = null
  let ctrl: AbortController | null = null
  onScopeDispose(() => ctrl?.abort())
  // An undo or a redo may have changed the names, which their outcome does not say.
  s.hooks.stepped = () => read()

  /** The source has names: the Name Box lists them, Formula AutoComplete offers them, the Name Manager shows them. */
  const available = computed(() => typeof s.props.source.names === 'function')
  /** Names can be defined and changed: an editable sheet whose source has `setNames` (and `names`, to show them). */
  const editable = computed(() => available.value && s.editable.value && typeof s.props.source.setNames === 'function')
  const collator = computed(() => new Intl.Collator(s.locale.value, { sensitivity: 'base' }))
  const compare = (a: string, b: string): number => collator.value.compare(a, b)

  /** Reads the names again; the previous list stays until the answer, and an earlier request's answer is dropped. */
  function read(): void {
    ctrl?.abort()
    ctrl = null
    const source = s.props.source
    if (readFrom !== source) {
      readFrom = source
      list.value = []
    }
    const call = s.sourceMethod('names')
    if (!call) return void (list.value = [])
    const outdated = s.outdated()
    const mine: AbortController = askSource(call, () => ctrl === mine && !outdated(), (got) => (list.value = got), () => {})
    ctrl = mine
  }

  /** The range each name refers to, null for a name whose formula is not a plain reference to a range (a constant, a formula, several areas). */
  const ranges = computed(() => {
    const out = new Map<SheetName, NameRange | null>()
    for (const n of list.value) {
      const place = parsePlace(n.ref)
      const sheet = place?.sheet ?? n.scope
      out.set(n, place?.range && sheet !== undefined ? { sheet, rect: { top: place.range.top, left: place.range.left, bottom: place.range.bottom, right: place.range.right } } : null)
    }
    return out
  })
  const rangeOf = (n: SheetName): NameRange | null => ranges.value.get(n) ?? null

  /** The names visible on a sheet, as Excel resolves them there: the sheet's own names, then the workbook's that none of them hides. */
  function visibleOn(sheet: string): SheetName[] {
    const own = list.value.filter((n) => n.scope !== undefined && sameName(n.scope, sheet))
    return [...own, ...list.value.filter((n) => n.scope === undefined && !own.some((o) => sameName(o.name, n.name)))]
  }
  /** The name typed in the Name Box or a formula on `sheet`: one visible there, or a sheet's own name typed with its sheet (`Sheet2!Total`). */
  function find(text: string, sheet: string): SheetName | undefined {
    const typed = text.trim()
    const scoped = SCOPED.exec(typed)
    if (scoped) {
      const on = scoped[1]?.replace(/''/g, "'") ?? scoped[2]!
      const found = list.value.find((n) => n.scope !== undefined && sameName(n.scope, on) && sameName(n.name, scoped[3]!))
      if (found) return found
    }
    return visibleOn(sheet).find((n) => sameName(n.name, typed))
  }
  /** The names visible on the active sheet in alphabetical order, for Formula AutoComplete. */
  const formulaNames = computed(() => visibleOn(s.sheetName.value).map((n) => n.name).sort(compare))
  /** What Excel's Name Box lists: the names visible on the active sheet that refer to ranges, in alphabetical order. */
  const boxNames = computed(() =>
    visibleOn(s.sheetName.value)
      .filter((n) => rangeOf(n) !== null)
      .sort((a, b) => compare(a.name, b.name)),
  )
  /** The name the Name Box shows for a selection on the active sheet: the first in its list whose range is exactly the selection. */
  function nameOf(rect: CellRect | null): string | null {
    if (!rect) return null
    const hit = boxNames.value.find((n) => {
      const r = rangeOf(n)!
      return sameName(r.sheet, s.sheetName.value) && r.rect.top === rect.top && r.rect.left === rect.left && r.rect.bottom === rect.bottom && r.rect.right === rect.right
    })
    return hit?.name ?? null
  }

  /**
   * Replaces the workbook's names through the source's `setNames`, as a single undo step; the sheet is read again afterwards
   * (a name changes the values of formulas anywhere) unless the source sent `reload` meanwhile, which reads the names too.
   * null when nothing was sent (no `setNames`, a read-only sheet); a refusal is returned for the caller to show.
   */
  async function save(next: SheetName[]): Promise<EditOutcome | null> {
    const call = editable.value ? s.sourceMethod('setNames') : null
    if (!call) return null
    const before = deps.reloads()
    const outdated = s.outdated()
    let out: EditOutcome
    try {
      out = await call(next)
    } catch (e) {
      out = failedOutcome(e)
    }
    if (outdated()) return out
    if (out.ok) deps.apply(out.changes)
    if ((out.ok || out.reason === 'conflict') && deps.reloads() === before) await deps.reload()
    return out
  }

  return { list, available, editable, compare, read, rangeOf, visibleOn, find, formulaNames, boxNames, nameOf, save }
}

export type Names = ReturnType<typeof useNames>
