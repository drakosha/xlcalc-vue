// Copy, cut and paste within the workbook, as in Excel: Ctrl+C and Ctrl+X put the selection's text on the clipboard and
// mark the range with a moving border; pasting that same text (or a clipboard the browser does not provide) pastes the
// range itself through the source's `structure` (`copyRange`, `moveRange`), with formulas adjusted as Excel adjusts them,
// and any other text is pasted as text. The cell menu does the same through the browser's Clipboard API. Excel's Paste
// Special pastes a copied range as all, formulas, values, formats or validation, transposed or not (`copyRange`'s `paste`
// and `transpose`, xlcalc 0.60).
import { computed, ref, shallowRef, watch, type Ref } from 'vue'
import { cellAddress, rangeAddress } from 'xlcalc/a1'
import type { CellRect, PasteKind } from 'xlcalc/contract'
import { formatClipboard } from 'xlcalc/text'
import type { SheetOp } from '../contract/source'
import type { Commands } from './commands'
import type { Editing, EditorState } from './editing'
import { onMac } from './keys'
import type { Selection } from './selection'
import type { SheetState } from './state'
import { listOf, type Totals } from './totals'
import type { Viewport } from './viewport'
import type { Windows } from './windows'

/** The range copied or cut on the grid, as Excel's copy mode holds it; `text` is what went to the clipboard, null while it is read. */
export interface ClipMark {
  sheet: string
  rect: CellRect
  text: string | null
  cut: boolean
  /** The displayed text the copy event put on the clipboard, when the source's own copy text replaced it there: pasting either counts as this copy. */
  shown?: string
}

export interface ClipDeps {
  notify: (text: string) => void
  editor: Ref<EditorState | null>
  /** Whether a column is visible: Excel omits hidden ones from a copy. */
  colShown: (col: number) => boolean
}

/** Clipboard text as compared: line endings of any system, without the last one. */
const plain = (t: string) => t.replace(/\r\n?/g, '\n').replace(/\n$/, '')

export function useClip(s: SheetState, w: Windows, sel: Selection, v: Viewport, ed: Editing, cmd: Commands, totals: Totals, deps: ClipDeps) {
  const mark = shallowRef<ClipMark | null>(null)
  /** The text of a copy mode that ended (Esc, an edit, a pasted cut): Excel has nothing to paste then, so pasting that text does nothing. */
  let spent: string | null = null
  function clear(): void {
    if (!mark.value) return
    spent = mark.value.text
    mark.value = null
  }
  // Excel's copy mode ends with a change of cells and when a cell editor opens.
  s.hooks.changing = clear
  watch(deps.editor, (e) => e && clear())

  const canCut = cmd.canRanges
  /** The menu's Paste: our copy, or a clipboard the browser may allow the page to read. */
  const canPaste = computed(() => s.editable.value && (!!mark.value || typeof navigator.clipboard?.readText === 'function'))
  /** Excel's Paste Special: over a copy (not a cut) in copy mode, with `structure` and no display filter. */
  const canPasteSpecial = computed(() => cmd.canRanges.value && !!mark.value && !mark.value.cut)
  const specialOpen = ref(false)
  /** Opens the Paste Special dialog; false when there is nothing it could paste. */
  function openSpecial(): boolean {
    if (!canPasteSpecial.value) return false
    specialOpen.value = true
    return true
  }
  /** The dialog's OK (with a choice) or Cancel: it closes, and the copy is pasted as chosen. */
  function closeSpecial(how?: { paste: PasteKind; transpose: boolean }): void {
    specialOpen.value = false
    s.focusGrid()
    const m = mark.value
    if (how && m && canPasteSpecial.value) void pasteMark(m, how)
  }
  /** Ctrl+Alt+V (Ctrl+Cmd+V on a Mac, as in Excel for Mac) opens Paste Special: true when handled. */
  function specialKey(e: KeyboardEvent): boolean {
    const keys = onMac() ? e.ctrlKey && e.metaKey && !e.altKey : e.ctrlKey && e.altKey && !e.metaKey
    if (!keys || e.shiftKey || e.code !== 'KeyV' || !canPasteSpecial.value) return false
    e.preventDefault()
    specialOpen.value = true
    return true
  }

  // ── the selection as text ──
  /** What a copy takes: the selection's display rows and columns, up to the data. */
  function extentOf(r: CellRect): { ds: number[]; c2: number } {
    return { ds: listOf(totals.shownRows(r.top, Math.min(r.bottom, Math.max(r.top, s.dataRows.value)))), c2: Math.min(r.right, Math.max(r.left, s.dataCols.value)) }
  }
  /** The visible rows and columns as Excel puts plain text on the clipboard (xlcalc's formatClipboard): tabs, CRLF, quotes around a tab or a line break. */
  const tsv = (cell: (d: number, col: number) => string, rows: number[], left: number, last: number) => {
    const cols: number[] = []
    for (let col = left; col <= last; col++) if (deps.colShown(col)) cols.push(col)
    return formatClipboard(
      rows.map((d) => cols.map((col) => cell(d, col))),
      { trailingNewline: !onMac() },
    )
  }
  /** The text from the loaded windows; null when they do not hold all of it. */
  function loadedText(r: CellRect): string | null {
    const { ds, c2 } = extentOf(r)
    if (!ds.length) return ''
    return totals.inWindows(ds, r.left, c2) ? tsv(w.textAt, ds, r.left, c2) : null
  }
  /** The text read from the source in one window, truncated by the cell limit. */
  async function fetchedText(r: CellRect): Promise<{ text: string; partial: boolean }> {
    const { ds, c2 } = extentOf(r)
    const d1 = ds[0]!
    const d2 = ds[ds.length - 1]!
    const filter = s.visible.value?.spec
    const got = await s.source().window({ sheet: s.sheetName.value, row: d1, rows: d2 - d1 + 1, col: r.left, cols: c2 - r.left + 1, ...(filter ? { filter } : {}) })
    if ('pending' in got) throw new Error('pending')
    const lastD = got.row + got.text.length - 1
    const lastCol = got.col + (got.text[0]?.length ?? 0) - 1
    const text = tsv((d, col) => got.text[d - got.row]?.[col - got.col] ?? '', ds.filter((d) => d <= lastD), r.left, Math.min(c2, lastCol))
    return { text, partial: lastD < d2 || lastCol < c2 }
  }

  // ── copy and cut ──
  /** Copies or cuts the selection: its text to the clipboard (`write` is the copy event's, otherwise the browser's Clipboard API) and the range marked. */
  async function put(cut: boolean, write: ((text: string) => void) | null): Promise<void> {
    const r = sel.selRect.value
    const full = sel.fullRect.value
    if (!r || !full) return
    // Under the display filter a copy is the text of the visible rows: there is no range to paste.
    const m: ClipMark | null = s.visible.value ? null : { sheet: s.sheetName.value, rect: full, text: null, cut }
    spent = null
    mark.value = m
    const now = loadedText(r)
    const exact = sourceText(r)
    if (now !== null && write) {
      write(now)
      named(m, now)
      if (exact) void replaceWith(mark.value, now, exact)
      return
    }
    // the source's copy, otherwise (none, or rejected as too large) the windows' text
    const fromWindows = async () => (now !== null ? { text: now, partial: false } : await fetchedText(r))
    let got: { text: string; partial: boolean }
    try {
      got = exact ? await exact.then((text) => ({ text, partial: false }), fromWindows) : await fromWindows()
    } catch {
      if (mark.value === m) mark.value = null
      return deps.notify(s.messages.value.copyFailed)
    }
    named(m, got.text)
    try {
      await navigator.clipboard.writeText(got.text)
      if (got.partial) deps.notify(s.messages.value.copyPartial)
    } catch {
      // The copy still pastes within the workbook.
      deps.notify(s.messages.value.clipboardDenied)
    }
  }
  /** The source's copy of the selection clipped to the data (values formatted by their number formats, as Excel copies them); null without `copyText`, under the display filter or beyond the data. */
  function sourceText(r: CellRect): Promise<string> | null {
    const copy = s.sourceMethod('copyText')
    const bottom = Math.min(r.bottom, Math.max(r.top, s.dataRows.value))
    const right = Math.min(r.right, Math.max(r.left, s.dataCols.value))
    if (!copy || s.visible.value || r.top > s.dataRows.value || r.left > s.dataCols.value) return null
    return copy(s.sheetName.value, rangeAddress({ top: r.top, left: r.left, bottom, right }), { trailingNewline: !onMac() })
  }
  /** Replaces the displayed text the copy event wrote with the source's text, through the Clipboard API (a browser that denies it keeps the displayed one); the mark knows both. */
  async function replaceWith(m: ClipMark | null, shown: string, exact: Promise<string>): Promise<void> {
    let text: string
    try {
      text = await exact
      if (text === shown || mark.value !== m) return
      await navigator.clipboard.writeText(text)
    } catch {
      return
    }
    if (m && mark.value === m) mark.value = { ...m, text, shown }
  }
  /** The mark takes its text once known, unless another copy happened meanwhile. */
  function named(m: ClipMark | null, text: string): void {
    if (m && mark.value === m) mark.value = { ...m, text }
  }
  function take(e: ClipboardEvent, cut: boolean): void {
    if (e.target !== s.scroller.value || !sel.selRect.value) return
    e.preventDefault()
    if (cut && !canCut.value) return
    const data = e.clipboardData
    void put(cut, data ? (text) => data.setData('text/plain', text) : null)
  }
  const onCopy = (e: ClipboardEvent) => take(e, false)
  const onCut = (e: ClipboardEvent) => take(e, true)

  // ── paste ──
  /** A paste on the grid or from the menu: the marked range when the clipboard holds its text or provides none (null), otherwise the text. */
  function pasteWith(text: string | null): void {
    if (!s.editable.value) return
    const m = mark.value
    if (m && (!text || m.text === null || plain(text) === plain(m.text) || (m.shown !== undefined && plain(text) === plain(m.shown)))) {
      if (cmd.canRanges.value) return void pasteMark(m)
      // Without `structure` (or under the display filter) the copy is pasted as text; its final CRLF ends the range rather than adding an empty row.
      text = m.text?.replace(/\r\n$/, '') ?? null
    } else if (!m && text && spent !== null && plain(text) === plain(spent)) return
    if (text) ed.pasteText(text)
  }
  /**
   * Pastes the marked range at the selection, as Excel pastes it: a copy is tiled over a selection whose size is a multiple
   * of it, otherwise pasted once from the top-left cell; a cut moves to the top-left cell and its mark ends. The pasted range is
   * selected. Paste Special's `how` pastes the copy as chosen, a transposed one once from the top-left cell, rows turned to columns.
   */
  async function pasteMark(m: ClipMark, how?: { paste: PasteKind; transpose: boolean }): Promise<void> {
    const a = sel.fullRect.value
    if (!a) return
    const sheet = s.sheetName.value
    const turned = !!how?.transpose
    const h = (turned ? m.rect.right - m.rect.left : m.rect.bottom - m.rect.top) + 1
    const wd = (turned ? m.rect.bottom - m.rect.top : m.rect.right - m.rect.left) + 1
    const tiles = !m.cut && !turned && (a.bottom - a.top + 1) % h === 0 && (a.right - a.left + 1) % wd === 0
    const target: CellRect = tiles ? a : { top: a.top, left: a.left, bottom: a.top + h - 1, right: a.left + wd - 1 }
    const lock = ed.lockIn(target) ?? (m.cut && m.sheet === sheet ? ed.lockIn(m.rect) : null)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const from = { sheet: m.sheet, range: rangeAddress(m.rect) }
    const corner = cellAddress(a.top, a.left)
    const to = { sheet, range: tiles ? rangeAddress(a) : corner }
    const op: SheetOp = m.cut ? { op: 'moveRange', from, to: { sheet, cell: corner } } : how ? { op: 'copyRange', from, to, paste: how.paste, ...(turned ? { transpose: true } : {}) } : { op: 'copyRange', from, to }
    const before = { selected: s.selected.value, extent: s.extent.value }
    const out = await cmd.structure(op, false)
    if (!out?.ok) return
    if (m.cut && mark.value === m) clear()
    if (sheet === s.sheetName.value && s.selected.value === before.selected && s.extent.value === before.extent) sel.selectRect(target)
  }
  function onPaste(e: ClipboardEvent): void {
    if (e.target !== s.scroller.value || !s.editable.value) return
    e.preventDefault()
    pasteWith(e.clipboardData ? e.clipboardData.getData('text/plain') : null)
  }

  // ── the cell menu ──
  /** Copy and Cut: the browser's Clipboard API takes the text; if it is denied, the copy still pastes within the workbook. */
  function copy(cut: boolean): void {
    if (cut && !canCut.value) return
    s.focusGrid()
    void put(cut, null)
  }
  /** Paste: the clipboard read through the browser's API; a browser that does not provide it pastes only our own copy. */
  async function paste(): Promise<void> {
    s.focusGrid()
    let text: string | null = null
    try {
      text = await navigator.clipboard.readText()
    } catch {
      text = null
    }
    if (text === null && !mark.value) return deps.notify(s.messages.value.pasteDenied)
    pasteWith(text)
  }

  /** The moving border around the marked range on its sheet. */
  const ants = computed(() => {
    const m = mark.value
    if (!m || m.sheet !== s.sheetName.value || s.visible.value) return null
    const bottom = Math.min(m.rect.bottom, s.rowCount.value)
    const right = Math.min(m.rect.right, s.colCount.value)
    return bottom < m.rect.top || right < m.rect.left ? null : v.boxOf(m.rect.top, m.rect.left, bottom, right)
  })

  return { mark, ants, canCut, canPaste, canPasteSpecial, specialOpen, openSpecial, closeSpecial, specialKey, clear, onCopy, onCut, onPaste, copy, paste }
}

export type Clip = ReturnType<typeof useClip>
