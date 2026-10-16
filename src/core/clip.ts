// Copy, cut and paste inside the book, as Excel's: Control+C and Control+X put the selection's text on the clipboard and
// mark the range with a moving border; a paste of that same text (or of a clipboard the browser will not read) pastes the
// range itself through the source's `structure` (`copyRange`, `moveRange`), formulas moved as Excel moves them, and any
// other text is pasted as text. The cells' menu does the same through the browser's clipboard API.
import { computed, shallowRef, watch, type Ref } from 'vue'
import { cellAddress, rangeAddress } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import { formatClipboard } from 'xlcalc/text'
import type { SheetOp } from '../source'
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
  /** The shown text the copy event took, when the source's own copy text replaced it on the clipboard: a paste of either is this copy. */
  shown?: string
}

export interface ClipDeps {
  notify: (text: string) => void
  editor: Ref<EditorState | null>
  /** Whether a column is shown: Excel leaves hidden ones out of a copy. */
  colShown: (col: number) => boolean
}

/** Clipboard text as compared: line ends of any system, without the last one. */
const plain = (t: string) => t.replace(/\r\n?/g, '\n').replace(/\n$/, '')

export function useClip(s: SheetState, w: Windows, sel: Selection, v: Viewport, ed: Editing, cmd: Commands, totals: Totals, deps: ClipDeps) {
  const mark = shallowRef<ClipMark | null>(null)
  /** The text of a copy mode that ended (Esc, an edit, a cut pasted): Excel has nothing to paste then, so that text is not pasted again. */
  let spent: string | null = null
  function clear(): void {
    if (!mark.value) return
    spent = mark.value.text
    mark.value = null
  }
  // Excel's copy mode ends with a change of cells and when a cell's field opens.
  s.hooks.changing = clear
  watch(deps.editor, (e) => e && clear())

  const canCut = computed(() => s.editable.value && cmd.can('structure') && !s.visible.value)
  /** The menu's Paste: our copy, or a clipboard the browser may let the page read. */
  const canPaste = computed(() => s.editable.value && (!!mark.value || typeof navigator.clipboard?.readText === 'function'))

  // ── the selection as text ──
  /** What a copy takes: display rows of the selection and its columns, up to the data. */
  function extentOf(r: CellRect): { ds: number[]; c2: number } {
    return { ds: listOf(totals.shownRows(r.top, Math.min(r.bottom, Math.max(r.top, s.dataRows.value)))), c2: Math.min(r.right, Math.max(r.left, s.dataCols.value)) }
  }
  /** The shown rows and columns as Excel puts plain text on the clipboard (xlcalc's formatClipboard): tabs, CRLF, quotes around a tab or a line break. */
  const tsv = (cell: (d: number, col: number) => string, rows: number[], left: number, last: number) => {
    const cols: number[] = []
    for (let col = left; col <= last; col++) if (deps.colShown(col)) cols.push(col)
    return formatClipboard(
      rows.map((d) => cols.map((col) => cell(d, col))),
      { trailingNewline: !onMac() },
    )
  }
  /** The text from the loaded windows; null when they do not hold it all. */
  function loadedText(r: CellRect): string | null {
    const { ds, c2 } = extentOf(r)
    if (!ds.length) return ''
    return totals.inWindows(ds, r.left, c2) ? tsv(w.textAt, ds, r.left, c2) : null
  }
  /** The text read from the source in one window, cut by the cell limit. */
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
  /** Copy or cut the selection: its text to the clipboard (`write` is the copy event's, else the browser's clipboard API) and the range marked. */
  async function put(cut: boolean, write: ((text: string) => void) | null): Promise<void> {
    const r = sel.selRect.value
    const full = sel.fullRect.value
    if (!r || !full) return
    // Under the display filter a copy is a text of the shown rows: there is no range to paste.
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
    // the source's copy, else (none, or refused as too large) the windows' text
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
      // The copy still pastes inside the book.
      deps.notify(s.messages.value.clipboardDenied)
    }
  }
  /** The source's copy of the selection cut to the data (values by their formats, as Excel copies them); null without `copyText`, under the display filter or past the data. */
  function sourceText(r: CellRect): Promise<string> | null {
    const copy = s.sourceMethod('copyText')
    const bottom = Math.min(r.bottom, Math.max(r.top, s.dataRows.value))
    const right = Math.min(r.right, Math.max(r.left, s.dataCols.value))
    if (!copy || s.visible.value || r.top > s.dataRows.value || r.left > s.dataCols.value) return null
    return copy(s.sheetName.value, rangeAddress({ top: r.top, left: r.left, bottom, right }), { trailingNewline: !onMac() })
  }
  /** The source's text in place of the shown one the copy event took, through the clipboard API (a browser that refuses keeps the shown one); the mark knows both. */
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
  /** The mark takes its text once known, unless another copy came meanwhile. */
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
  /** A paste on the grid or from the menu: the marked range when the clipboard holds its text or gives none (null), else the text. */
  function pasteWith(text: string | null): void {
    if (!s.editable.value) return
    const m = mark.value
    if (m && (!text || m.text === null || plain(text) === plain(m.text) || (m.shown !== undefined && plain(text) === plain(m.shown)))) {
      if (cmd.can('structure') && !s.visible.value) return void pasteMark(m)
      // Without `structure` (or under the display filter) the copy goes as its text; its last CRLF ends the rectangle, not an empty row to paste.
      text = m.text?.replace(/\r\n$/, '') ?? null
    } else if (!m && text && spent !== null && plain(text) === plain(spent)) return
    if (text) ed.pasteText(text)
  }
  /**
   * The marked range pasted at the selection, as Excel pastes it: a copy tiles a selection its size divides, else goes
   * once from the top-left cell; a cut moves to the top-left cell and its mark ends. The pasted range is selected.
   */
  async function pasteMark(m: ClipMark): Promise<void> {
    const a = sel.fullRect.value
    if (!a) return
    const sheet = s.sheetName.value
    const h = m.rect.bottom - m.rect.top + 1
    const wd = m.rect.right - m.rect.left + 1
    const tiles = !m.cut && (a.bottom - a.top + 1) % h === 0 && (a.right - a.left + 1) % wd === 0
    const target: CellRect = tiles ? a : { top: a.top, left: a.left, bottom: a.top + h - 1, right: a.left + wd - 1 }
    const lock = ed.lockIn(target) ?? (m.cut && m.sheet === sheet ? ed.lockIn(m.rect) : null)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const from = { sheet: m.sheet, range: rangeAddress(m.rect) }
    const corner = cellAddress(a.top, a.left)
    const op: SheetOp = m.cut ? { op: 'moveRange', from, to: { sheet, cell: corner } } : { op: 'copyRange', from, to: { sheet, range: tiles ? rangeAddress(a) : corner } }
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

  // ── the cells' menu ──
  /** Copy and Cut: the browser's clipboard API takes the text; refused, the copy still pastes inside the book. */
  function copy(cut: boolean): void {
    if (cut && !canCut.value) return
    s.focusGrid()
    void put(cut, null)
  }
  /** Paste: the clipboard read through the browser's API; a browser that will not give it pastes only our own copy. */
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

  /** The moving border round the marked range on its sheet. */
  const ants = computed(() => {
    const m = mark.value
    if (!m || m.sheet !== s.sheetName.value || s.visible.value) return null
    const bottom = Math.min(m.rect.bottom, s.rowCount.value)
    const right = Math.min(m.rect.right, s.colCount.value)
    return bottom < m.rect.top || right < m.rect.left ? null : v.boxOf(m.rect.top, m.rect.left, bottom, right)
  })

  return { mark, ants, canCut, canPaste, clear, onCopy, onCut, onPaste, copy, paste }
}

export type Clip = ReturnType<typeof useClip>
