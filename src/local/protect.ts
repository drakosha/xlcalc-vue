// Excel's sheet protection in the local source: the session does not enforce it, so each change that reaches a protected
// sheet is checked first and refused as Excel refuses it, with Excel's alert in the workbook's language. What Excel allows
// was measured in Excel for Mac 16.114 through its object model (docs/local.md, Protection).
import type { AsyncSession, BookLayout, CellRect, CellStyle, SheetLayout, StructureOp } from 'xlcalc'
import { cellAddress, colNumber, MAX_COLS, MAX_ROWS, parseCell, parseLines, parseRange, rangeAddress, sameName } from 'xlcalc/a1'
import { parseClipboard } from 'xlcalc/text'
import type { CellEntry, EditOutcome, FindRequest, FormatPatch, LayoutChange, OutlineChange, SheetOp } from '../contract/source'
import type { ClearKind, SheetRules, TableOp } from '../contract/workbook'
import { allows, lockedStyle, type ProtectionAllowance } from '../protection'
import { boxesUnder, replacedCells } from './find'
import { isPivotOp, pivotSheets } from './pivots'

/** The outcome code of a change a protected sheet does not allow. */
export const PROTECTED = 'protected'

/** Excel's alerts in English and in Russian: the protected sheet's, as Excel for Windows words it (also for deleting a row or
 * a column with a locked cell, which has no text of its own there), and Excel for Mac's for a table's command. */
const TEXTS = {
  en: {
    alert: "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.",
    tables: "Table features aren't available because the sheet is protected.",
  },
  ru: {
    alert: 'Ячейка или диаграмма, которую вы пытаетесь изменить, находится на защищенном листе. Чтобы внести изменения, снимите защиту листа. Может потребоваться пароль.',
    tables: 'Функции таблицы невозможно использовать на защищенном листе.',
  },
}

type Refusal = Extract<EditOutcome, { ok: false }>
/** A change to check: the cells it writes (each must be unlocked) and the boxes it needs, on one sheet; `refuse` when nothing allows it. */
interface Need {
  sheet: string
  cells?: CellRect[]
  /** Scattered cells (Replace All's): each box's listed cells. */
  boxes?: { rect: CellRect; only: Set<string> }[]
  allow?: ProtectionAllowance[]
  refuse?: true
  message?: string
}

/** A sheet's rows or columns that have a style of their own, from its layout. */
function styledLines(map: Record<string, { style?: number }> | undefined, axis: 'rows' | 'columns'): { first: number; last: number; style: number }[] {
  const out: { first: number; last: number; style: number }[] = []
  for (const [key, line] of Object.entries(map ?? {})) {
    const span = line.style === undefined ? null : parseLines(key, axis)
    if (span) out.push({ first: span.first, last: span.last, style: line.style! })
  }
  return out
}

/** How many of `first`..`last` a span covers. */
const overlap = (first: number, last: number, span: { first: number; last: number }) => Math.max(0, Math.min(last, span.last) - Math.max(first, span.first) + 1)

/**
 * Whether a part of the sheet beyond its used range holds a locked cell: there a cell takes its row's style, else its
 * column's, else none (locked).
 */
function emptyPartLocked(r: CellRect, sheet: SheetLayout, styles: readonly CellStyle[]): boolean {
  let styledRows = 0
  for (const span of styledLines(sheet.rows, 'rows')) {
    const n = overlap(r.top, r.bottom, span)
    if (!n) continue
    if (lockedStyle(styles[span.style])) return true
    styledRows += n
  }
  if (styledRows === r.bottom - r.top + 1) return false
  let styledCols = 0
  for (const span of styledLines(sheet.columns, 'columns')) {
    const n = overlap(r.left, r.right, span)
    if (!n) continue
    if (lockedStyle(styles[span.style])) return true
    styledCols += n
  }
  return styledCols < r.right - r.left + 1
}

/**
 * Whether any cell of `rect` (of `only` among them) is locked: within the used range by the viewport's resolved styles, read
 * in windows of `maxCells`, beyond it by rows and columns.
 */
async function lockedIn(s: AsyncSession, sheet: SheetLayout, styles: readonly CellStyle[], rect: CellRect, maxCells: number, only?: Set<string>): Promise<boolean> {
  const r = { top: rect.top, left: rect.left, bottom: Math.min(MAX_ROWS, rect.bottom), right: Math.min(MAX_COLS, rect.right) }
  const used = await s.getUsedRange(sheet.name)
  const inBottom = Math.min(r.bottom, used.rows)
  const inRight = Math.min(r.right, used.columns)
  if (inBottom >= r.top && inRight >= r.left) {
    const width = Math.min(inRight - r.left + 1, maxCells)
    const height = Math.max(1, Math.floor(maxCells / width))
    for (let left = r.left; left <= inRight; left += width) {
      const right = Math.min(inRight, left + width - 1)
      for (let top = r.top; top <= inBottom; top += height) {
        const bottom = Math.min(inBottom, top + height - 1)
        const v = await s.getViewport(sheet.name, rangeAddress({ top, left, bottom, right }), { maxCells })
        for (let row = top; row <= bottom; row++) {
          for (let col = left; col <= right; col++) {
            const addr = cellAddress(row, col)
            if (only && !only.has(addr)) continue
            const id = v.cells[addr]?.effectiveStyle
            if (lockedStyle(id === undefined ? undefined : v.styles[id])) return true
          }
        }
      }
    }
  }
  if (r.bottom > used.rows && emptyPartLocked({ ...r, top: Math.max(r.top, used.rows + 1) }, sheet, styles)) return true
  return r.right > used.columns && inBottom >= r.top && emptyPartLocked({ ...r, bottom: inBottom, left: Math.max(r.left, used.columns + 1) }, sheet, styles)
}

/** The rectangle a paste of clipboard text at a cell fills. */
function pasteRect(cell: string, text: string): CellRect | null {
  const at = parseCell(cell)
  if (!at) return null
  const rows = parseClipboard(text)
  const cols = Math.max(1, ...rows.map((r) => r.length))
  return { top: at.row, left: at.col, bottom: at.row + Math.max(1, rows.length) - 1, right: at.col + cols - 1 }
}

/** Where a copy lands: its target, or from the target's top-left cell the copy's size (turned for a transpose), whichever is larger. */
function copyTarget(from: CellRect, to: CellRect, transpose: boolean | undefined): CellRect {
  const h = transpose ? from.right - from.left + 1 : from.bottom - from.top + 1
  const w = transpose ? from.bottom - from.top + 1 : from.right - from.left + 1
  return { top: to.top, left: to.left, bottom: Math.max(to.bottom, to.top + h - 1), right: Math.max(to.right, to.left + w - 1) }
}

/** Entire rows or columns as a rectangle. */
const linesRect = (axis: 'rows' | 'columns', first: number, count: number): CellRect =>
  axis === 'rows' ? { top: first, left: 1, bottom: first + count - 1, right: MAX_COLS } : { top: 1, left: first, bottom: MAX_ROWS, right: first + count - 1 }

/** The sheet protection of a local source over its session, in the workbook's `locale`, reading windows of `maxCells`. */
export function protectionGuard(locale: string, maxCells: number) {
  const words = locale.toLowerCase().startsWith('ru') ? TEXTS.ru : TEXTS.en
  const refusal = (message = words.alert): Refusal => ({ ok: false, reason: 'refused', code: PROTECTED, message })
  /** The workbook's layout, read once until the next change. */
  let layout: Promise<BookLayout> | null = null
  const layoutOf = (s: AsyncSession): Promise<BookLayout> => (layout ??= s.getLayout())
  /** A change was made or tried (the source's every change ends with it): the layout is read again before the next check. */
  const forget = () => void (layout = null)

  /** Why a protected sheet refuses `needs`, or null when every sheet allows them. */
  async function check(s: AsyncSession, needs: Need[]): Promise<Refusal | null> {
    const book = await layoutOf(s)
    for (const need of needs) {
      const sheet = book.sheets.find((x) => sameName(x.name, need.sheet))
      const p = sheet?.protection
      if (!sheet || !p) continue
      if (need.refuse || need.allow?.some((a) => !allows(p, a))) return refusal(need.message)
      for (const rect of need.cells ?? []) if (await lockedIn(s, sheet, book.styles ?? [], rect, maxCells)) return refusal(need.message)
      for (const b of need.boxes ?? []) if (await lockedIn(s, sheet, book.styles ?? [], b.rect, maxCells, b.only)) return refusal(need.message)
    }
    return null
  }
  /** Cells by address, a range or entire lines (`B2`, `B2:D9`, `B:D`, `2:5`) as rectangles; one that is none is left to the session. */
  const rects = (ranges: string[]): CellRect[] => ranges.flatMap((x) => parseRange(x) ?? [])

  /** What a structure operation needs; a table's sheet is found by its name. */
  async function structureNeeds(s: AsyncSession, op: StructureOp): Promise<Need[]> {
    // Use PivotTable and PivotChart: "format, change the layout, refresh, or otherwise modify PivotTables, or create new reports" (Microsoft's Protect a worksheet)
    if (isPivotOp(op)) return pivotSheets(await layoutOf(s), op).map((sheet) => ({ sheet, allow: ['pivotTables'] }))
    switch (op.op) {
      case 'insertRows':
        return [{ sheet: op.sheet, allow: ['insertRows'] }]
      case 'insertColumns':
        return [{ sheet: op.sheet, allow: ['insertColumns'] }]
      case 'deleteRows':
        return [{ sheet: op.sheet, allow: ['deleteRows'], cells: [linesRect('rows', op.row, op.count ?? 1)] }]
      case 'deleteColumns': {
        const col = colNumber(op.column) ?? 1
        return [{ sheet: op.sheet, allow: ['deleteColumns'], cells: [linesRect('columns', col, op.count ?? 1)] }]
      }
      case 'moveRange': {
        const from = parseRange(op.from.range)
        const at = parseCell(op.to.cell)
        if (!from || !at) return []
        const to = { top: at.row, left: at.col, bottom: at.row + from.bottom - from.top, right: at.col + from.right - from.left }
        return [{ sheet: op.from.sheet, cells: [from] }, { sheet: op.to.sheet, cells: [to] }]
      }
      case 'copyRange': {
        const from = parseRange(op.from.range)
        const to = parseRange(op.to.range)
        if (!from || !to) return []
        const cells = [copyTarget(from, to, op.transpose)]
        // Paste Special's parts that Excel gates by the protection's boxes rather than by the cells
        if (op.paste === 'formats') return [{ sheet: op.to.sheet, allow: ['formatCells'], cells }]
        if (op.paste === 'columnWidths') return [{ sheet: op.to.sheet, allow: ['formatColumns'] }]
        if (op.paste === 'notes') return [{ sheet: op.to.sheet, allow: ['objects'] }]
        if (op.paste === 'validation') return [{ sheet: op.to.sheet, refuse: true }]
        return [{ sheet: op.to.sheet, cells }]
      }
      case 'fillRange':
        return [{ sheet: op.sheet, cells: rects([op.to]) }]
      case 'mergeCells':
      case 'unmergeCells':
      case 'group':
      case 'ungroup':
      case 'showDetail':
      case 'hideDetail':
      case 'showLevels':
        return [{ sheet: op.sheet, refuse: true }]
      case 'addTable':
        return [{ sheet: op.sheet, refuse: true, message: words.tables }]
      case 'renameTable':
      case 'deleteTable':
      case 'convertTable': {
        const sheet = (await layoutOf(s)).sheets.find((x) => x.tables?.some((t) => sameName(t.name, op.table)))
        return sheet ? [{ sheet: sheet.name, refuse: true, message: words.tables }] : []
      }
      default:
        // sheets themselves are the workbook's structure, which a sheet's protection does not cover
        return []
    }
  }

  /** A method's check: given the method, it runs the method once `needs` are allowed, otherwise answers with the refusal (`refused`). */
  function guard<P extends unknown[]>(needs: (s: AsyncSession, ...args: P) => Need[] | Promise<Need[]>) {
    return <A extends [...P, ...unknown[]], R>(run: (s: AsyncSession, ...args: A) => Promise<R>, refused: (r: Refusal) => R = (r) => r as R) =>
      async (s: AsyncSession, ...args: A): Promise<R> => {
        const no = await check(s, await needs(s, ...(args as unknown as P)))
        return no ? refused(no) : run(s, ...args)
      }
  }

  /** Each source method with the protection Excel checks before it. */
  return {
    forget,
    edit: guard((_s, entries: CellEntry[]) => entries.map((e) => ({ sheet: e.sheet, cells: rects([e.cell]) }))),
    paste: guard((_s, sheet: string, cell: string, text: string) => {
      const r = pasteRect(cell, text)
      return r ? [{ sheet, cells: [r] }] : []
    }),
    fill: guard((_s, sheet: string, range: string) => [{ sheet, cells: rects([range]) }]),
    pick: guard((_s, sheet: string, cell: string) => [{ sheet, cells: rects([cell]) }]),
    structure: guard((s, op: SheetOp) => structureNeeds(s, op)),
    /** Why a protected sheet refuses an operation the source makes itself (Refresh All's of the PivotTables); null when it allows it. */
    refusal: async (s: AsyncSession, op: StructureOp): Promise<Refusal | null> => check(s, await structureNeeds(s, op)),
    table: guard((s, op: TableOp) => structureNeeds(s, op)),
    // Excel's Group and Ungroup are disabled on a protected sheet, and its outline's buttons refuse
    outline: guard((_s, change: OutlineChange) => [{ sheet: change.sheet, refuse: true }]),
    layout: guard((_s, change: LayoutChange) => [{ sheet: change.sheet, allow: ['columns' in change ? 'formatColumns' : 'formatRows'] }]),
    /** Formatting needs Format cells, over locked cells too; Locked and Hidden cannot change on a protected sheet, nor can a cell's own style go. */
    format: guard((_s, sheet: string, patches: FormatPatch[]) => [
      { sheet, allow: ['formatCells'] },
      ...(patches.some((p) => p.style === null || (p.style && 'protection' in p.style)) ? [{ sheet, refuse: true as const }] : []),
    ]),
    clear: guard((_s, sheet: string, ranges: string[]) => [{ sheet, cells: rects(ranges) }]),
    /** Data > Filter is disabled on a protected sheet whatever its boxes allow: they let users use an AutoFilter, not set or remove one. */
    setAutoFilter: guard((_s, sheet: string) => [{ sheet, refuse: true }]),
    link: guard((_s, sheet: string, range: string) => [{ sheet, allow: ['insertHyperlinks'], cells: rects([range]) }]),
    note: guard((_s, sheet: string) => [{ sheet, allow: ['objects'] }]),
    /** Conditional formats are Format cells' (Excel's Format Cells and Conditional Formatting dialogs); Data Validation is disabled. */
    setRules: guard((_s, sheet: string, rules: Partial<SheetRules>) => [{ sheet, allow: rules.conditionalFormats ? ['formatCells'] : [] }, ...(rules.validations ? [{ sheet, refuse: true as const }] : [])]),
    /** Clear All over unlocked cells without Format cells clears their contents and keeps their formats, as Excel did. */
    async clearKind(s: AsyncSession, sheet: string, what: ClearKind): Promise<ClearKind | 'contents'> {
      if (what !== 'all') return what
      const p = (await layoutOf(s)).sheets.find((x) => sameName(x.name, sheet))?.protection
      return allows(p, 'formatCells') ? 'all' : 'contents'
    },
    /** Replace All: every cell Find finds must be unlocked. */
    replace: <R>(run: (s: AsyncSession, req: FindRequest, by: string) => Promise<R>, refused: (r: Refusal) => R, rows: (s: AsyncSession, req: FindRequest) => Promise<number[] | null>) =>
      guard(async (s, req: FindRequest) => {
        const p = (await layoutOf(s)).sheets.find((x) => sameName(x.name, req.sheet))?.protection
        if (!p) return []
        return [{ sheet: req.sheet, boxes: boxesUnder(req.sheet, await replacedCells(s, req, await rows(s, req)), maxCells) }]
      })(run, refused),
  }
}

export type ProtectionGuard = ReturnType<typeof protectionGuard>

/** Excel's Protect Sheet and Unprotect Sheet: a single `setLayout` step; the sheets are read again. */
export async function protectSheet(s: AsyncSession, sheet: string, protection: SheetLayout['protection'] | null, done: (changes: Awaited<ReturnType<AsyncSession['setLayout']>>['changes']) => EditOutcome): Promise<EditOutcome> {
  return done((await s.setLayout([{ sheet, protection: protection ?? null }])).changes)
}
