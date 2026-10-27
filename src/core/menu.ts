// Excel's context menus: on row numbers and column letters (insert, delete, sizes, hide, unhide, group, ungroup), on the
// sheet tabs (insert, delete, rename, move, hide, unhide) and the "+" after them, on cells (cut, copy, paste, the display
// filter's Reapply, Freeze Panes and its variants, Trace Precedents, then the notes' and hyperlinks' items); a right click
// opens one, and so does a long press on a touch screen. Its small dialogs ask for a size, confirm a sheet's deletion,
// list hidden sheets, and ask whether to group rows or columns.
import { computed, nextTick, onScopeDispose, ref, shallowRef } from 'vue'
import { colLetters, linesKey, MAX_COLS, MAX_ROWS, rangeAddress, rectContains, sameName } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import { columnWidthFromChars, columnWidthToChars, MAX_ROW_HEIGHT } from 'xlcalc/text'
import { linesOn } from '../lines'
import { cellRect } from '../rect'
import type { SheetInfo, SheetOp } from '../contract/source'
import type { Clip } from './clip'
import type { Commands } from './commands'
import type { Explain } from './explain'
import type { Filter } from './filter'
import { ptFromPx, widthFromPx } from './geometry'
import type { Outline, OutlineAxis, OutlineOp } from './outline'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** A finger held this long on a row number, a column letter or a tab opens the menu; moved beyond the slop, it scrolls instead. */
const LONG_PRESS_MS = 500
const PRESS_SLOP_PX = 10
/** The release, click and browser context menu that follow a long press belong to it for this long. */
const PRESS_TAIL_MS = 800
/** Excel's limit of the Column Width box, in characters; the Row Height box's limit is the tallest row, `MAX_ROW_HEIGHT`. */
const MAX_WIDTH_CHARS = 255

export type MenuKind = 'rows' | 'columns' | 'sheet' | 'cells'

export interface MenuItem {
  id: string
  label: string
  run: () => void
  disabled?: boolean
  /** A separator above it, between groups. */
  group?: boolean
}

/** The open menu: where it is and what it acts on (the selected rows or columns, a sheet by its index; the cell menu acts on the selection). */
export interface OpenMenu {
  kind: MenuKind
  x: number
  y: number
  first: number
  last: number
}

/** A small dialog over the sheet: a size in Excel's units, the confirmation before a sheet with data is deleted, the hidden sheets, rows or columns to group. */
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
  /** Whether the workbook freezes panes on the displayed sheet: the first freeze item is then Unfreeze Panes, as in Excel. */
  frozen: () => boolean
  /** Excel's Define Name in the cell menu: offered where names can be defined. */
  defineName?: { available: () => boolean; run: () => void }
  /** The cell menu's items for notes and for hyperlinks, which Excel lists before and after Define Name at the menu's end. */
  cellNotes: () => MenuItem[]
  cellLinks: () => MenuItem[]
}

export function useMenu(s: SheetState, sel: Selection, cmd: Commands, deps: MenuDeps) {
  const open = shallowRef<OpenMenu | null>(null)
  const dialog = ref<MenuDialog | null>(null)
  /** A tab being renamed in place: its sheet's index and the draft. */
  const renaming = ref<{ index: number; draft: string } | null>(null)
  const close = () => (open.value = null)
  const m = () => s.messages.value

  // ── what the menu acts on ──
  const linesAvailable = computed(() => cmd.can('structure') || cmd.can('layout') || cmd.can('outline'))
  /** Opens the menu of a row number or a column letter: a row or column outside a selection of entire rows or columns is selected first, as in Excel. */
  function openLines(kind: 'rows' | 'columns', line: number, x: number, y: number): boolean {
    if (!linesAvailable.value) return false
    const whole = kind === 'rows' ? sel.wholeRows.value : sel.wholeColumns.value
    // Under a display filter the rows between two visible ones are not visible rows: the menu acts on the clicked row.
    const lines = sel.linesOf(kind === 'rows' && s.visible.value ? null : whole, line)
    if (lines !== whole) (kind === 'rows' ? sel.pickRow : sel.pickColumn)(line)
    open.value = { kind, x, y, first: lines.first, last: Math.min(lines.last, kind === 'rows' ? MAX_ROWS : MAX_COLS) }
    return true
  }
  /** Opens a tab's menu: the sheet becomes active, as a right click on a tab makes it in Excel. */
  async function openSheet(index: number, x: number, y: number): Promise<boolean> {
    if (!cmd.can('structure') || !s.sheets.value[index]) return false
    await deps.pickSheet(index)
    open.value = { kind: 'sheet', x, y, first: index, last: index }
    return true
  }

  /** Opens the cell menu: a cell outside the selection is selected first, as in Excel. */
  function openCells(row: number, col: number, x: number, y: number): void {
    const r = sel.fullRect.value
    if (!r || !rectContains(r, cellRect(row, col))) sel.pickCell(row, col)
    open.value = { kind: 'cells', x, y, first: 0, last: 0 }
  }

  const items = computed<MenuItem[]>(() => {
    const o = open.value
    if (!o) return []
    if (o.kind === 'cells') return cellItems()
    return o.kind === 'sheet' ? sheetItems(o.first) : lineItems(o.kind, o.first, o.last)
  })
  /** Cut, Copy and Paste within the workbook (a read-only sheet only copies), Reapply under a display filter, Freeze Panes with the source's `freeze`, Trace Precedents with its `explain`, Define Name with its `setNames`. */
  function cellItems(): MenuItem[] {
    const clip = deps.clip
    const copy: MenuItem = { id: 'copy', label: m().menuCopy, run: () => clip.copy(false) }
    const out: MenuItem[] = !s.editable.value
      ? [copy]
      : [
          { id: 'cut', label: m().menuCut, disabled: !clip.canCut.value, run: () => clip.copy(true) },
          copy,
          { id: 'paste', label: m().menuPaste, disabled: !clip.canPaste.value, run: () => void clip.paste() },
          { id: 'pasteSpecial', label: m().menuPasteSpecial, disabled: !clip.canPasteSpecial.value, run: () => void clip.openSpecial() },
        ]
    if (deps.filter.active.value && deps.filter.available.value) out.push({ id: 'reapply', label: m().menuReapply, run: () => void deps.filter.reapply(), group: true })
    // the display filter's rows on screen are not the sheet's rows, which panes freeze
    if (cmd.can('freeze') && !s.visible.value) {
      const frozen = deps.frozen()
      out.push(
        { id: frozen ? 'unfreeze' : 'freeze', label: frozen ? m().menuUnfreezePanes : m().menuFreezePanes, run: () => void cmd.freeze(frozen ? 'unfreeze' : 'panes'), group: true },
        { id: 'freezeTopRow', label: m().menuFreezeTopRow, run: () => void cmd.freeze('topRow') },
        { id: 'freezeFirstColumn', label: m().menuFreezeFirstColumn, run: () => void cmd.freeze('firstColumn') },
      )
    }
    if (deps.explain.available.value) out.push({ id: 'precedents', label: m().menuPrecedents, run: () => deps.explain.show(), group: true })
    out.push(...deps.cellNotes())
    const define = deps.defineName
    if (define?.available()) out.push({ id: 'defineName', label: m().menuDefineName, run: define.run, group: true })
    out.push(...deps.cellLinks())
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
      const lines = linesOn(axis, key)
      out.push(
        { id: 'size', label: axis === 'rows' ? m().menuRowHeight : m().menuColumnWidth, run: () => openSize(axis, first, last), group: out.length > 0 },
        { id: 'hide', label: m().menuHide, run: () => void cmd.layout({ sheet, ...lines, hidden: true }) },
        // always present, as in Excel: the grid knows a hidden row or column only from the windows it has read
        { id: 'unhide', label: m().menuUnhide, run: () => void cmd.layout({ sheet, ...lines, hidden: false }) },
      )
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
      // Excel moves a sheet among the visible ones: before the visible sheet to its left, or after the one to its right.
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
  /** A column's width as the workbook stores it: one dragged meanwhile, from its pixels, otherwise the workbook's; null for the default. */
  const storedWidth = (col: number): number | null => {
    const px = s.userWidths.value[col]
    return px !== undefined ? widthFromPx(px) : (s.bookWidths.value.get(col) ?? null)
  }
  function openSize(axis: 'rows' | 'columns', first: number, last: number): void {
    const size = axis === 'columns' ? columnWidthToChars(storedWidth(first)) : ptFromPx(s.heights.value.get(first) ?? s.defaultRowPx.value)
    dialog.value = { kind: 'size', axis, first, last, draft: decimals(size), error: '' }
  }
  /** The dialog's number as typed, with a decimal comma or point; NaN for anything else. */
  const parseSize = (text: string): number => (/^\s*\d+(?:[.,]\d*)?\s*$/.test(text) ? Number(text.trim().replace(',', '.')) : Number.NaN)
  function applySize(): void {
    const d = dialog.value
    if (d?.kind !== 'size') return
    const n = parseSize(d.draft)
    const max = d.axis === 'columns' ? MAX_WIDTH_CHARS : MAX_ROW_HEIGHT
    if (!(n >= 0 && n <= max)) {
      d.error = d.axis === 'columns' ? m().widthRange : m().heightRange
      return
    }
    dialog.value = null
    const sheet = s.sheetName.value
    const key = linesKey(d.first, d.last, d.axis)
    const lines = linesOn(d.axis, key)
    // A size of 0 hides the rows or columns, as Excel's dialogs accept it; their previous size comes back when they are unhidden.
    if (n === 0) void cmd.layout({ sheet, ...lines, hidden: true })
    else if (d.axis === 'columns') void cmd.layout({ sheet, columns: key, width: columnWidthFromChars(n) })
    else void cmd.layout({ sheet, rows: key, height: n })
  }

  // ── Group and Ungroup over a cell range: Excel's dialog asks whether to group rows or columns, rows first ──
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
  /** The last number this sheet gave a new sheet: Excel keeps counting past one deleted since. */
  let issued = 0
  /** Excel's name for a new sheet: the interface's word and the next number beyond the sheet count and the numbers already in names. */
  function nextSheetName(): string {
    const prefix = m().sheetNamePrefix
    const own = new RegExp(`^${prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\d+)$`, 'i')
    let n = Math.max(issued, s.sheets.value.length, ...s.sheets.value.map((x) => Number(own.exec(x.name)?.[1] ?? 0)))
    do n++
    while (s.sheets.value.some((x) => sameName(x.name, `${prefix}${n}`)))
    issued = n
    return `${prefix}${n}`
  }
  /** "+": a new sheet at the end, displayed once it is added. */
  const addSheet = () => cmd.structure({ op: 'insertSheet', name: nextSheetName() })
  /** Excel asks before deleting a sheet with data; an empty one is deleted at once. */
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

  // ── renaming on the tab, as in Excel: Enter keeps the name, Esc discards it, a rejection keeps the input ──
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
  /** After a rejection Enter keeps the input, so the name can be fixed; leaving the input then discards it. */
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

  // A long press on a touch screen, shared by row numbers, column letters, cells, tabs and the formula bar's references: what follows it belongs to each.
  const press = useLongPress()

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
    addSheet,
    confirmDelete,
    unhide,
    startRename,
    commitRename,
    cancelRename,
    press,
  }
}

export type Menu = ReturnType<typeof useMenu>

/** A long press on a touch screen: held still for `LONG_PRESS_MS`, it opens what `start` was given; moved beyond the slop, it scrolls instead. */
export function useLongPress() {
  let press: { timer: ReturnType<typeof setTimeout>; x: number; y: number } | null = null
  let pressedAt = -Infinity
  /** A finger down: `open` runs at its point if it stays still long enough. */
  function start(e: PointerEvent, open: (x: number, y: number) => void): void {
    if (e.pointerType !== 'touch' || e.isPrimary === false) return
    cancel()
    const { clientX: x, clientY: y } = e
    press = {
      x,
      y,
      timer: setTimeout(() => {
        press = null
        pressedAt = Date.now()
        open(x, y)
      }, LONG_PRESS_MS),
    }
  }
  function move(e: PointerEvent): void {
    if (press && Math.hypot(e.clientX - press.x, e.clientY - press.y) > PRESS_SLOP_PX) cancel()
  }
  function cancel(): void {
    if (press) clearTimeout(press.timer)
    press = null
  }
  /** Whether a click or a context menu now is what follows a long press that opened something. */
  const after = () => Date.now() - pressedAt < PRESS_TAIL_MS
  /** Whether a finger is down and still, its long press yet to come. */
  const pressing = () => press !== null
  /** The browser's context menu where ours opens: suppresses the browser's own and cancels a press in progress; false for what follows a long press that opened ours (a browser sends it afterwards), true otherwise. */
  function context(e: MouseEvent): boolean {
    e.preventDefault()
    cancel()
    return !after()
  }
  // The sheet is unmounted: a press held then opens nothing.
  onScopeDispose(cancel)
  return { start, move, cancel, after, pressing, context }
}
