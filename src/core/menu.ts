// The menus as Excel's: on row numbers and column letters (insert, delete, sizes, hide, unhide, group, ungroup), on the
// sheet tabs (insert, delete, rename, move, hide, unhide) and "+" after them, on cells (cut, copy, paste, the display
// filter's Reapply, Freeze Panes and its kinds, Trace Precedents); a right click
// opens one, a finger's long press too. Its small dialogs ask a size, confirm a sheet's deletion, list hidden sheets, ask
// rows or columns.
import { computed, nextTick, ref, shallowRef } from 'vue'
import { colLetters, linesKey, MAX_COLS, MAX_ROWS, rangeAddress, rectContains, sameName } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import type { SheetInfo, SheetOp } from '../source'
import type { Clip } from './clip'
import type { Commands } from './commands'
import type { Explain } from './explain'
import type { Filter } from './filter'
import { charsOfPx, ptFromPx, widthFromChars, type Geometry } from './geometry'
import type { Outline, OutlineAxis, OutlineOp } from './outline'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** A finger held this long on a number, a letter or a tab opens the menu; moved further than the slop it scrolls instead. */
const LONG_PRESS_MS = 500
const PRESS_SLOP_PX = 10
/** The lift, click and browser's context menu that follow a long press are its own for this long. */
const PRESS_TAIL_MS = 800
/** Excel's limits of the Column Width (characters) and Row Height (points) boxes. */
const MAX_WIDTH_CHARS = 255
const MAX_HEIGHT_PT = 409

export type MenuKind = 'rows' | 'columns' | 'sheet' | 'cells'

export interface MenuItem {
  id: string
  label: string
  run: () => void
  disabled?: boolean
  /** A line above it, between groups. */
  group?: boolean
}

/** The open menu: where it is, what it acts on (lines of the selection, a sheet by its index; the cells' menu acts on the selection). */
export interface OpenMenu {
  kind: MenuKind
  x: number
  y: number
  first: number
  last: number
}

/** A small dialog over the sheet: a size in Excel's units, the question before a sheet with data goes, the hidden sheets, rows or columns to group. */
export type MenuDialog =
  | { kind: 'size'; axis: 'columns' | 'rows'; first: number; last: number; draft: string; error: string }
  | { kind: 'deleteSheet'; sheet: string }
  | { kind: 'unhide'; sheets: string[] }
  | { kind: 'outline'; op: OutlineOp; axis: OutlineAxis; rect: CellRect }

export interface MenuDeps {
  pickSheet: (index: number) => Promise<void>
  tabs: () => { sheet: SheetInfo; i: number }[]
  outline: Outline
  clip: Clip
  filter: Filter
  explain: Explain
  /** Whether the book freezes panes on the sheet shown: the first freeze item is then Unfreeze Panes, as Excel's. */
  frozen: () => boolean
}

export function useMenu(s: SheetState, geo: Geometry, sel: Selection, cmd: Commands, deps: MenuDeps) {
  const open = shallowRef<OpenMenu | null>(null)
  const dialog = ref<MenuDialog | null>(null)
  /** A tab being renamed in place: its sheet's index and the draft. */
  const renaming = ref<{ index: number; draft: string } | null>(null)
  const close = () => (open.value = null)
  const m = () => s.messages.value

  // ── what the menu acts on ──
  const linesAvailable = computed(() => cmd.can('structure') || cmd.can('layout') || cmd.can('outline'))
  /** Open the menu of a row number or a column letter: a line outside a selection of whole lines selects it first, as Excel. */
  function openLines(kind: 'rows' | 'columns', line: number, x: number, y: number): boolean {
    if (!linesAvailable.value) return false
    const whole = kind === 'rows' ? sel.wholeRows.value : sel.wholeColumns.value
    // Under a display filter the rows between two shown ones are not the shown ones: the menu acts on the row pressed.
    const inside = whole && line >= whole.first && line <= whole.last && !(kind === 'rows' && s.visible.value)
    if (!inside) (kind === 'rows' ? sel.pickRow : sel.pickColumn)(line)
    const lines = (kind === 'rows' ? sel.wholeRows.value : sel.wholeColumns.value)!
    open.value = { kind, x, y, first: lines.first, last: Math.min(lines.last, kind === 'rows' ? MAX_ROWS : MAX_COLS) }
    return true
  }
  /** Open a tab's menu: the sheet becomes the active one, as a right click on a tab makes it in Excel. */
  async function openSheet(index: number, x: number, y: number): Promise<boolean> {
    if (!cmd.can('structure') || !s.sheets.value[index]) return false
    await deps.pickSheet(index)
    open.value = { kind: 'sheet', x, y, first: index, last: index }
    return true
  }

  /** Open the cells' menu: a cell outside the selection selects it first, as Excel. */
  function openCells(row: number, col: number, x: number, y: number): void {
    const r = sel.fullRect.value
    if (!r || !rectContains(r, { top: row, left: col, bottom: row, right: col })) sel.pickCell(row, col)
    open.value = { kind: 'cells', x, y, first: 0, last: 0 }
  }

  /** Whether lines first..last hold a hidden one the grid knows of (a size of 0). */
  function hiddenIn(axis: 'rows' | 'columns', first: number, last: number): boolean {
    if (axis === 'columns') {
      for (let c = first; c <= Math.min(last, s.colCount.value); c++) if (geo.widthOf(c) === 0) return true
      return false
    }
    for (const [r, px] of s.heights.value) if (px === 0 && r >= first && r <= last) return true
    return false
  }

  const items = computed<MenuItem[]>(() => {
    const o = open.value
    if (!o) return []
    if (o.kind === 'cells') return cellItems()
    return o.kind === 'sheet' ? sheetItems(o.first) : lineItems(o.kind, o.first, o.last)
  })
  /** Cut, Copy and Paste inside the book (a read-only sheet only copies), Reapply under a display filter, Freeze Panes with the source's `freeze`, Trace Precedents with its `explain`. */
  function cellItems(): MenuItem[] {
    const clip = deps.clip
    const copy: MenuItem = { id: 'copy', label: m().menuCopy, run: () => clip.copy(false) }
    const out: MenuItem[] = !s.editable.value
      ? [copy]
      : [
          { id: 'cut', label: m().menuCut, disabled: !clip.canCut.value, run: () => clip.copy(true) },
          copy,
          { id: 'paste', label: m().menuPaste, disabled: !clip.canPaste.value, run: () => void clip.paste() },
        ]
    if (deps.filter.active.value && deps.filter.available.value) out.push({ id: 'reapply', label: m().menuReapply, run: () => void deps.filter.reapply(), group: true })
    // the display filter's rows on screen are not the sheet's, which panes freeze
    if (cmd.can('freeze') && !s.visible.value) {
      const frozen = deps.frozen()
      out.push(
        { id: frozen ? 'unfreeze' : 'freeze', label: frozen ? m().menuUnfreezePanes : m().menuFreezePanes, run: () => void cmd.freeze(frozen ? 'unfreeze' : 'panes'), group: true },
        { id: 'freezeTopRow', label: m().menuFreezeTopRow, run: () => void cmd.freeze('topRow') },
        { id: 'freezeFirstColumn', label: m().menuFreezeFirstColumn, run: () => void cmd.freeze('firstColumn') },
      )
    }
    if (deps.explain.available.value) out.push({ id: 'precedents', label: m().menuPrecedents, run: () => deps.explain.show(), group: true })
    return out
  }
  function lineItems(axis: 'rows' | 'columns', first: number, last: number): MenuItem[] {
    const sheet = s.sheetName.value
    const count = last - first + 1
    const key = linesKey(first, last, axis)
    const out: MenuItem[] = []
    if (cmd.can('structure')) {
      const column = colLetters(first)
      const insert: SheetOp = axis === 'rows' ? { op: 'insertRows', sheet, row: first, count } : { op: 'insertColumns', sheet, column, count }
      const remove: SheetOp = axis === 'rows' ? { op: 'deleteRows', sheet, row: first, count } : { op: 'deleteColumns', sheet, column, count }
      out.push({ id: 'insert', label: m().menuInsert, run: () => void cmd.structure(insert) }, { id: 'delete', label: m().menuDelete, run: () => void cmd.structure(remove) })
    }
    if (cmd.can('layout')) {
      const lines = axis === 'rows' ? { rows: key } : { columns: key }
      out.push(
        { id: 'size', label: axis === 'rows' ? m().menuRowHeight : m().menuColumnWidth, run: () => openSize(axis, first, last), group: out.length > 0 },
        { id: 'hide', label: m().menuHide, run: () => void cmd.layout({ sheet, ...lines, hidden: true }) },
      )
      if (hiddenIn(axis, first, last)) out.push({ id: 'unhide', label: m().menuUnhide, run: () => void cmd.layout({ sheet, ...lines, hidden: false }) })
    }
    if (cmd.can('outline')) {
      out.push(
        { id: 'group', label: m().menuGroup, run: () => deps.outline.lines('group', axis, first, last), group: out.length > 0 },
        { id: 'ungroup', label: m().menuUngroup, run: () => deps.outline.lines('ungroup', axis, first, last) },
      )
    }
    return out
  }
  function sheetItems(index: number): MenuItem[] {
    const info = s.sheets.value[index]
    if (!info) return []
    const name = info.name
    const tabs = deps.tabs()
    const k = tabs.findIndex((t) => t.i === index)
    const hidden = s.sheets.value.filter((x) => x.hidden === true).map((x) => x.name)
    const out: MenuItem[] = [
      { id: 'insertSheet', label: m().menuInsert, run: () => void cmd.structure({ op: 'insertSheet', name: nextSheetName(), before: name }) },
      { id: 'deleteSheet', label: m().menuDelete, run: () => deleteSheet(info) },
      { id: 'rename', label: m().menuRename, run: () => startRename(index) },
      // Excel moves a sheet among the shown ones: before the shown sheet left of it, or past the one right of it.
      { id: 'moveLeft', label: m().menuMoveLeft, disabled: k <= 0, run: () => void cmd.structure({ op: 'moveSheet', sheet: name, before: tabs[k - 1]!.sheet.name }), group: true },
      { id: 'moveRight', label: m().menuMoveRight, disabled: k < 0 || k >= tabs.length - 1, run: () => void cmd.structure(k + 2 < tabs.length ? { op: 'moveSheet', sheet: name, before: tabs[k + 2]!.sheet.name } : { op: 'moveSheet', sheet: name }) },
      { id: 'hideSheet', label: m().menuHide, run: () => void cmd.structure({ op: 'hideSheet', sheet: name }), group: true },
    ]
    if (hidden.length) out.push({ id: 'unhideSheet', label: m().menuUnhideSheet, run: () => (dialog.value = { kind: 'unhide', sheets: hidden }) })
    return out
  }
  function choose(item: MenuItem): void {
    if (item.disabled) return
    close()
    item.run()
  }

  // ── sizes in Excel's units: characters for a width, points for a height ──
  const decimals = (x: number) => new Intl.NumberFormat(s.locale.value, { maximumFractionDigits: 2, useGrouping: false }).format(x)
  function openSize(axis: 'rows' | 'columns', first: number, last: number): void {
    const size = axis === 'columns' ? charsOfPx(geo.widthOf(first)) : ptFromPx(s.heights.value.get(first) ?? s.defaultRowPx.value)
    dialog.value = { kind: 'size', axis, first, last, draft: decimals(size), error: '' }
  }
  /** The box's number as typed, a decimal comma or point; NaN for anything else. */
  const parseSize = (text: string): number => (/^\s*\d+(?:[.,]\d*)?\s*$/.test(text) ? Number(text.trim().replace(',', '.')) : Number.NaN)
  function applySize(): void {
    const d = dialog.value
    if (d?.kind !== 'size') return
    const n = parseSize(d.draft)
    const max = d.axis === 'columns' ? MAX_WIDTH_CHARS : MAX_HEIGHT_PT
    if (!(n >= 0 && n <= max)) {
      d.error = d.axis === 'columns' ? m().widthRange : m().heightRange
      return
    }
    dialog.value = null
    const sheet = s.sheetName.value
    const key = linesKey(d.first, d.last, d.axis)
    const lines = d.axis === 'rows' ? { rows: key } : { columns: key }
    // A size of 0 hides the lines, as Excel's boxes take it; the size they had comes back when they are shown.
    if (n === 0) void cmd.layout({ sheet, ...lines, hidden: true })
    else if (d.axis === 'columns') void cmd.layout({ sheet, columns: key, width: widthFromChars(n) })
    else void cmd.layout({ sheet, rows: key, height: n })
  }

  // ── Group and Ungroup over a selection of cells: Excel's box asks rows or columns, rows first ──
  function askOutline(op: OutlineOp): void {
    const rect = sel.fullRect.value
    if (rect) dialog.value = { kind: 'outline', op, axis: 'rows', rect }
  }
  function applyOutline(): void {
    const d = dialog.value
    if (d?.kind !== 'outline') return
    dialog.value = null
    if (d.axis === 'rows') deps.outline.lines(d.op, 'rows', d.rect.top, Math.min(d.rect.bottom, MAX_ROWS))
    else deps.outline.lines(d.op, 'columns', d.rect.left, Math.min(d.rect.right, MAX_COLS))
  }

  // ── sheets ──
  /** The last number this sheet gave a new sheet: Excel counts on past one deleted since. */
  let issued = 0
  /** Excel's name for a new sheet: the interface's word and the next number past the sheets' count and their own numbers. */
  function nextSheetName(): string {
    const prefix = m().sheetNamePrefix
    const own = new RegExp(`^${prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\d+)$`, 'i')
    let n = Math.max(issued, s.sheets.value.length, ...s.sheets.value.map((x) => Number(own.exec(x.name)?.[1] ?? 0)))
    do n++
    while (s.sheets.value.some((x) => sameName(x.name, `${prefix}${n}`)))
    issued = n
    return `${prefix}${n}`
  }
  /** "+": a new sheet at the end, shown once it is in. */
  const addSheet = () => cmd.structure({ op: 'insertSheet', name: nextSheetName() })
  /** Excel asks before a sheet with data goes; an empty one goes at once. */
  function deleteSheet(info: SheetInfo): void {
    if (info.maxRow > 0 && info.maxCol > 0) dialog.value = { kind: 'deleteSheet', sheet: info.name }
    else void cmd.structure({ op: 'deleteSheet', sheet: info.name })
  }
  function confirmDelete(): void {
    const d = dialog.value
    if (d?.kind !== 'deleteSheet') return
    dialog.value = null
    void cmd.structure({ op: 'deleteSheet', sheet: d.sheet })
  }
  function unhide(name: string): void {
    dialog.value = null
    void cmd.structure({ op: 'unhideSheet', sheet: name })
  }

  // ── renaming on the tab, as Excel's: Enter keeps the name, Esc drops it, a refusal keeps the field ──
  const renameInput = ref<HTMLInputElement | null>(null)
  function startRename(index: number): void {
    if (!cmd.can('structure') || !s.sheets.value[index]) return
    renaming.value = { index, draft: s.sheets.value[index]!.name }
    void nextTick(() => {
      renameInput.value?.focus()
      renameInput.value?.select()
    })
  }
  let renameBusy = false
  /** Enter keeps the field after a refusal, for the name to be mended; leaving the field drops it then. */
  async function commitRename(left = false): Promise<void> {
    const r = renaming.value
    if (!r || renameBusy) return
    const old = s.sheets.value[r.index]?.name
    if (old === undefined || r.draft === '' || r.draft === old) return void (renaming.value = null)
    renameBusy = true
    try {
      const out = await cmd.structure({ op: 'renameSheet', sheet: old, name: r.draft })
      if (out?.ok || left) renaming.value = null
      else void nextTick(() => renameInput.value?.focus())
    } finally {
      renameBusy = false
    }
  }
  const cancelRename = () => (renaming.value = null)

  // ── a finger's long press ──
  let press: { timer: ReturnType<typeof setTimeout>; x: number; y: number } | null = null
  let pressedAt = -Infinity
  /** A finger down: the menu opens if it stays put long enough. */
  function pressStart(e: PointerEvent, openAt: (x: number, y: number) => void): void {
    if (e.pointerType !== 'touch' || e.isPrimary === false) return
    pressCancel()
    const { clientX: x, clientY: y } = e
    press = {
      x,
      y,
      timer: setTimeout(() => {
        press = null
        pressedAt = Date.now()
        openAt(x, y)
      }, LONG_PRESS_MS),
    }
  }
  function pressMove(e: PointerEvent): void {
    if (press && Math.hypot(e.clientX - press.x, e.clientY - press.y) > PRESS_SLOP_PX) pressCancel()
  }
  function pressCancel(): void {
    if (press) clearTimeout(press.timer)
    press = null
  }
  /** Whether a click or a context menu now is the tail of a long press that opened the menu. */
  const afterPress = () => Date.now() - pressedAt < PRESS_TAIL_MS

  /** The range the open menu acts on, for its caption (`2:5`, `B:D`, `B2:C4`, the sheet's name). */
  const caption = computed(() => {
    const o = open.value
    if (!o) return ''
    if (o.kind === 'sheet') return s.sheets.value[o.first]?.name ?? ''
    if (o.kind === 'cells') return sel.selectionRef.value
    return rangeAddress(o.kind === 'rows' ? { top: o.first, left: 1, bottom: o.last, right: MAX_COLS } : { top: 1, left: o.first, bottom: MAX_ROWS, right: o.last })
  })

  return {
    open,
    items,
    caption,
    dialog,
    renaming,
    renameInput,
    close,
    choose,
    openLines,
    openSheet,
    openCells,
    linesAvailable,
    applySize,
    askOutline,
    applyOutline,
    nextSheetName,
    addSheet,
    confirmDelete,
    unhide,
    startRename,
    commitRename,
    cancelRename,
    pressStart,
    pressMove,
    pressCancel,
    afterPress,
  }
}

export type Menu = ReturnType<typeof useMenu>
