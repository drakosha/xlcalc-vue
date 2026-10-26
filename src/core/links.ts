// Excel's Insert Link dialog (Ctrl+K) and the cell menu's hyperlink items: Link, Edit Hyperlink, Open Hyperlink, Remove Hyperlink.
// A link goes to the source's `link`; a changed Text to Display goes to `edit` before it, as a separate undo step.
import { computed, ref } from 'vue'
import { cellAddress, quoteSheetIfNeeded, rangeAddress, sameName } from 'xlcalc/a1'
import type { SheetLinkEvent } from '../contract/view'
import type { SheetWindow } from '../contract/source'
import type { SheetLink } from '../contract/workbook'
import { safeHref } from './cell-view'
import type { Editing } from './editing'
import type { Extras } from './extras'
import type { Formula } from './formula'
import type { MenuItem } from './menu'
import { parsePlace } from './place'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** Excel's Link to: an address outside the workbook, or a place in it. */
export type LinkTab = 'file' | 'place'

/** A place in this document the dialog lists: a sheet, with the reference typed for it, or a defined name, which is the place itself. */
export interface LinkPlace {
  key: string
  label: string
}
/** The dialog's list under its headings: Cell Reference (the sheets) and Defined Names. */
export interface LinkPlaceGroup {
  kind: 'sheets' | 'names'
  label: string
  places: LinkPlace[]
}

/** The open dialog: what it links (the selection), its fields, and the link it edits (Edit Hyperlink). */
export interface LinkDraft {
  sheet: string
  range: string
  cell: string
  /** The link of the active cell when there is one: the dialog is then Edit Hyperlink, with Remove Link. */
  old: SheetLink | null
  tab: LinkTab
  /** Text to Display: editable for a single cell without a formula; `null` shows Excel's `<<Selection in Document>>`. */
  text: string | null
  /** The cell's displayed text when the dialog opened: the text is written only when it changed. */
  shown: string
  /** The text was typed in: until then, for an empty cell, it follows the address as Excel's does. */
  textTyped: boolean
  address: string
  ref: string
  place: string
  tooltip: string
  error: string
}

export interface LinksDeps {
  notify: (text: string) => void
  /** Whether the selection is the active cell alone, or the merged range it is in. */
  single: () => boolean
  /** Goes to a location or a defined name, as a click on a hyperlink does. */
  goPlace: (text: string) => Promise<boolean>
  /** The workbook's defined names that refer to ranges, for Place in This Document. */
  names: () => string[]
  /** Whether the host opens links through `openLink`, and the event to it. */
  hostLinks: () => boolean
  openLink: (link: SheetLinkEvent) => void
}

/** A place key: a sheet's (with the reference typed for it) or a defined name's. */
const sheetKey = (name: string) => `sheet:${name}`
const nameKey = (name: string) => `name:${name}`

export function useLinks(s: SheetState, sel: Selection, w: Windows, f: Formula, ed: Editing, extras: Extras, deps: LinksDeps) {
  const m = () => s.messages.value
  const draft = ref<LinkDraft | null>(null)
  const available = computed(() => extras.method('link') !== null)

  /** The active cell's hyperlink from its loaded window. */
  function activeLink(): NonNullable<SheetWindow['links']>[string] | undefined {
    const a = s.selected.value
    return a ? w.linkOfRow(a.row, a.col) : undefined
  }

  /**
   * The places Excel's dialog lists: the workbook's visible sheets under Cell Reference, and under Defined Names the
   * workbook's names that refer to ranges, with a location the edited link has that is not a cell reference.
   */
  const places = computed<LinkPlaceGroup[]>(() => {
    const sheets = s.sheets.value.filter((x) => !x.hidden).map((x) => ({ key: sheetKey(x.name), label: x.name }))
    const groups: LinkPlaceGroup[] = [{ kind: 'sheets', label: m().linkCellReference, places: sheets }]
    const old = draft.value?.old?.location
    const listed = deps.names()
    const extra = old && !parsePlace(old) && !listed.some((n) => sameName(n, old)) ? [old] : []
    const names = [...listed, ...extra].map((n) => ({ key: nameKey(n), label: n }))
    if (names.length) groups.push({ kind: 'names', label: m().linkDefinedNames, places: names })
    return groups
  })
  /** Whether the chosen place is a sheet, whose cell reference is typed; a defined name takes none, as Excel greys the box out. */
  const refOn = computed(() => !!draft.value?.place.startsWith('sheet:'))

  /** The location of the place chosen: `'Sheet 2'!B12` from a sheet and its typed reference, or a defined name; null for a reference that is not one. */
  function locationOf(d: LinkDraft): string | null {
    if (d.place.startsWith('name:')) return d.place.slice(5)
    const typed = parsePlace(d.ref)
    if (!typed?.range || typed.sheet !== undefined) return null
    return `${quoteSheetIfNeeded(d.place.slice(6))}!${rangeAddress(typed.range)}`
  }
  /** What the link shows in an empty cell: its address, or its place as typed. */
  function linkText(d: LinkDraft): string {
    if (d.tab === 'file') return d.address.trim()
    return d.place.startsWith('name:') ? d.place.slice(5) : `${d.place.slice(6)}!${d.ref.trim()}`
  }
  /** Text to Display follows the address in an empty cell until it is typed in, as Excel's does. */
  function follow(): void {
    const d = draft.value
    if (d && d.text !== null && !d.textTyped && d.shown === '') d.text = linkText(d)
  }

  /** Excel's Ctrl+K: Insert Link over the selection, or Edit Hyperlink when the active cell has a link; false when links cannot be edited. */
  function open(): boolean {
    const a = s.selected.value
    if (!available.value || !s.sheet.value || !a || extras.locked(sel.selRect.value)) return false
    const old = activeLink()
    const at = old?.location ? parsePlace(old.location.replace(/^#/, '')) : null
    const atSheet = at?.sheet !== undefined ? s.sheets.value[f.sheetIndexByName(at.sheet)]?.name : undefined
    const place = old?.location && !at ? nameKey(old.location) : sheetKey(atSheet ?? s.sheetName.value)
    const shown = w.textOfRow(a.row, a.col)
    const own = sel.multi.value && !deps.single() ? null : f.formulaAt(a.row, a.col) === undefined ? shown : null
    draft.value = {
      sheet: s.sheetName.value,
      range: sel.selectionRef.value,
      cell: cellAddress(a.row, a.col),
      old: old ? { ...(old.href ? { href: old.href } : {}), ...(old.location ? { location: old.location } : {}), ...(old.tooltip ? { tooltip: old.tooltip } : {}) } : null,
      tab: old && !old.href && old.location ? 'place' : 'file',
      text: own,
      shown,
      textTyped: false,
      address: old?.href ?? '',
      ref: at?.range ? rangeAddress(at.range) : 'A1',
      place,
      tooltip: old?.tooltip ?? '',
      error: '',
    }
    return true
  }
  function close(): void {
    draft.value = null
    s.focusGrid()
  }

  /** OK: the link from the chosen tab; a changed Text to Display is entered first, as its own undo step. */
  async function apply(): Promise<void> {
    const d = draft.value
    const call = extras.method('link')
    if (!d || !call) return
    let link: SheetLink
    if (d.tab === 'file') {
      const href = d.address.trim()
      if (!href) return close()
      // a place the edited link had beside its address stays with it, as Excel keeps a file's subaddress
      link = { href, ...(d.old?.href && d.old.location ? { location: d.old.location } : {}) }
    } else {
      const location = locationOf(d)
      if (location === null) return void (d.error = m().linkBadRef)
      link = { location }
    }
    const tip = d.tooltip.trim()
    if (tip) link.tooltip = tip
    close()
    if (d.text !== null) {
      const text = d.text === '' ? linkText({ ...d, text: null }) : d.text
      if (text !== d.shown) {
        const out = await ed.send([{ sheet: d.sheet, cell: d.cell, text }], null)
        if (!out?.ok) return
      }
    }
    await extras.change(() => call(d.sheet, d.range, link), d.sheet, [d.range])
  }

  /** Remove Link in the dialog, and Remove Hyperlink in the cell menu: the links of the range go, the cells' contents stay. */
  async function remove(range = draft.value?.range ?? sel.selectionRef.value): Promise<void> {
    const call = extras.method('link')
    const sheet = draft.value?.sheet ?? s.sheetName.value
    if (draft.value) close()
    if (!call || !range || extras.locked(sel.selRect.value)) return
    await extras.change(() => call(sheet, range, null), sheet, [range])
  }

  /**
   * Follows a cell's hyperlink as a click on it does: a place in the workbook is selected, anything else goes to the host when
   * it listens, otherwise a web link opens in a new tab; false when the link is left to the browser's own click.
   */
  function openLink(link: NonNullable<SheetWindow['links']>[string], row: number, col: number, viaClick = false): boolean {
    if (link.location) {
      void deps.goPlace(link.location.replace(/^#/, ''))
      return true
    }
    if (deps.hostLinks()) {
      deps.openLink({ ...link, sheet: s.sheetName.value, cell: cellAddress(row, col) })
      return true
    }
    if (viaClick) return false
    if (safeHref(link.href)) window.open(link.href!.trim(), '_blank', 'noopener')
    return true
  }

  /** The cell menu's hyperlink items, as Excel's: Link on a cell without one; Edit, Open and Remove Hyperlink on a cell with one (Open alone on a read-only sheet). */
  function items(): MenuItem[] {
    const link = activeLink()
    const a = s.selected.value
    const out: MenuItem[] = []
    if (available.value && !link) out.push({ id: 'link', label: m().menuLink, run: () => void open() })
    if (link && a) {
      if (available.value) out.push({ id: 'editLink', label: m().menuEditLink, run: () => void open() })
      out.push({ id: 'openLink', label: m().menuOpenLink, run: () => void openLink(link, a.row, a.col) })
      if (available.value) out.push({ id: 'removeLink', label: m().menuRemoveLink, run: () => void remove(sel.selectionRef.value) })
    }
    if (out.length) out[0]!.group = true
    return out
  }

  return { draft, available, places, refOn, follow, open, close, apply, remove, openLink, items }
}

export type Links = ReturnType<typeof useLinks>
