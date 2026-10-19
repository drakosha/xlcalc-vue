// One drawn cell: its classes, CSS and tooltip, the content box, borders shared with neighbors, merges and text overflow.
import { computed, type ComputedRef, type Ref } from 'vue'
import { cellAddress, rectContains } from 'xlcalc/a1'
import type { BorderLine, CellStyle, Icon } from 'xlcalc/contract'
import { cellRect } from '../rect'
import { alignCss, borderCss, borderWidth, decorationCss, fillText, fontCss, horizontalOf, overflows, setSide, verticalOf, wrapsOf, type CssLine, type Horizontal } from './cell-style'
import { barLayers, type Layer } from './conditional'
import type { Editing, EditorState } from './editing'
import type { Formula, Highlight } from './formula'
import type { Geometry } from './geometry'
import type { Navigation } from './navigation'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Viewport } from './viewport'
import type { CellData, Windows } from './windows'

/** A cell's vertical padding; with its borders it fits inside the row's height. */
const CELL_PAD_V = 1
/** How far text may overflow into empty neighbors, in pixels. */
const OVERFLOW_MAX = 2000

/** A merged range as drawn: its size in the workbook (box height) and in drawn rows and columns (hidden ones are not in the DOM). */
export interface MergeSpan {
  rs: number
  cs: number
  rsDom: number
  csDom: number
}

export type ContentKind = 'text' | 'fill' | 'overflow' | 'href' | 'place' | 'host' | 'dead' | 'icon'

export interface CellView {
  attrs: { class: string[]; style: Record<string, string>; title: string; rowspan?: number; colspan?: number; 'data-cell': string }
  box: { class: string[]; style: Record<string, string> }
  kind: ContentKind
  text: string
  /** Underline and strikethrough, set on the element that holds the text. */
  deco: Record<string, string>
  /** Overflow into empty neighbors: the pixels taken on the left and right, the text's position, its background (hides crossed gridlines, as in Excel). */
  ov?: { left: number; right: number; justify: string; bg: string }
  link?: CellData['link']
  /** A conditional format's icon, drawn at the left of the cell before its text. */
  icon?: Icon
  /** The draft typed in the formula bar replaces `text`, drawn separately so typing does not redraw the grid; `fill`: the width Excel's Fill alignment repeats it across. */
  draft?: { fill?: number }
}

/** The frequently changing state that a drawn row's cells need, computed per row: a change redraws only the rows it affects. */
export interface RowState {
  /** The row's height in pixels. */
  height: number
  /** The selection over the row: its columns, whether the row is its top, how many rows below this one it ends (counted only as far as the row's merges reach), several cells or one; null when it does not cross the row. */
  sel: { left: number; right: number; top: boolean; below: number; multi: boolean } | null
  /** The active cell's column in the row; 0 for none. */
  active: number
  /** The formula's ranges over the row. */
  hl: readonly Highlight[]
  /** Columns of the search matches in the row, and of the current one (0: not in the row). */
  hits: ReadonlySet<number> | undefined
  hitNow: number
  /** The columns of the range just jumped to, over the row. */
  flash: { left: number; right: number } | null
  /** The editor open on a cell of the row: its column, and whether it is the formula bar's. */
  editor: { col: number; bar: boolean } | null
}

/** Whether two states of a row draw the same. */
export function sameRowState(a: RowState, b: RowState): boolean {
  const same = <T extends object>(x: T | null, y: T | null) => x === y || (!!x && !!y && (Object.keys(x) as (keyof T)[]).every((k) => x[k] === y[k]))
  return (
    a.height === b.height &&
    a.active === b.active &&
    a.hits === b.hits &&
    a.hitNow === b.hitNow &&
    a.hl.length === b.hl.length &&
    a.hl.every((h, i) => h === b.hl[i]) &&
    same(a.sel, b.sel) &&
    same(a.flash, b.flash) &&
    same(a.editor, b.editor)
  )
}

/** Whether two maps of drawn merges hold the same spans. */
function sameMerges(a: Map<string, MergeSpan | null>, b: Map<string, MergeSpan | null>): boolean {
  if (a.size !== b.size) return false
  for (const [key, x] of b) {
    const y = a.get(key)
    if (x === null ? y !== null : !y || y.rs !== x.rs || y.cs !== x.cs || y.rsDom !== x.rsDom || y.csDom !== x.csDom) return false
  }
  return true
}

/** Links a browser may open on its own; anything else (`javascript:` and the like) goes to the host or is only displayed. */
export const safeHref = (href: string | undefined): boolean => !!href && /^(https?:|mailto:|tel:|ftp:)/i.test(href.trim())

interface CellViewDeps {
  editor: Ref<EditorState | null>
  /** Search matches by `row:col`, and the current one. */
  find: { hitSet: ComputedRef<Set<string>>; current: ComputedRef<string> }
  /** Whether the host listens to `openLink`. */
  hostLinks: () => boolean
  /** List mode's `openRows`: the first column's cells open their rows. */
  openRows: () => boolean
}

export function useCellView(
  s: SheetState,
  geo: Geometry,
  v: Viewport,
  w: Windows,
  sel: Selection,
  f: Formula,
  nav: Navigation,
  ed: Editing,
  deps: CellViewDeps,
) {
  /** Merges of the drawn area by sheet rows: the top-left cell's span, null for covered cells. A merge not drawn in full (or crossing the freeze edge) is drawn as plain cells; under a filter only one-row merges stay. */
  const merges = computed((prev?: Map<string, MergeSpan | null>) => {
    const map = new Map<string, MergeSpan | null>()
    const rowsOk = (d1: number, d2: number) => d2 <= v.frozenRows.value || (d1 >= v.firstRow.value && d2 <= v.lastRow.value)
    const colsOk = (c1: number, c2: number) => c2 <= v.frozenCols.value || (c1 >= v.firstCol.value && c2 <= v.lastCol.value)
    for (const m of w.merges.value) {
      if (s.visible.value && m.top !== m.bottom) continue
      const d1 = s.indexOf(m.top)
      if (d1 === undefined) continue
      const d2 = d1 + m.bottom - m.top
      if (!rowsOk(d1, d2) || !colsOk(m.left, m.right)) continue
      let rsDom = 0
      let csDom = 0
      for (let d = d1; d <= d2; d++) if (geo.heightOf(d) > 0) rsDom++
      for (let c = m.left; c <= m.right; c++) if (geo.widthOf(c) > 0) csDom++
      if (!rsDom || !csDom) continue
      // The top-left cell may be hidden: the merge then starts at its first visible cell.
      const d0 = geo.heightOf(d1) > 0 ? d1 : geo.nextRow(d1, 1)
      const c0 = geo.widthOf(m.left) > 0 ? m.left : geo.nextCol(m.left, 1)
      for (let r = m.top; r <= m.bottom; r++) for (let c = m.left; c <= m.right; c++) map.set(`${r}:${c}`, null)
      map.set(`${s.rowOf(d0)}:${c0}`, { rs: d2 - d0 + 1, cs: m.right - c0 + 1, rsDom, csDom })
    }
    // The same map while the drawn merges are the same: every cell reads it, and sizes and scrolling recompute it.
    return prev && sameMerges(prev, map) ? prev : map
  })
  /** A cell's merge as drawn; no key is built for an area without merges (the grid asks for every drawn cell). */
  const spanOf = (r: number, c: number) => (merges.value.size ? merges.value.get(`${r}:${c}`) : undefined)
  const coveredByMerge = (r: number, c: number) => spanOf(r, c) === null

  /** The most rows a merge drawn from a row spans, by sheet row: a selection's bottom edge matters to a row only that far. */
  const mergeRows = computed(() => {
    const map = new Map<number, number>()
    for (const [key, span] of merges.value) {
      if (span && span.rs > 1) {
        const r = Number(key.slice(0, key.indexOf(':')))
        map.set(r, Math.max(map.get(r) ?? 1, span.rs))
      }
    }
    return map
  })
  /** The search matches by sheet row: their columns. */
  const hitRows = computed(() => {
    const map = new Map<number, Set<number>>()
    for (const key of deps.find.hitSet.value) {
      const at = key.indexOf(':')
      const r = Number(key.slice(0, at))
      let cols = map.get(r)
      if (!cols) map.set(r, (cols = new Set()))
      cols.add(Number(key.slice(at + 1)))
    }
    return map
  })
  /** The frequently changing state that a row's cells read, computed for display row `d` and sheet row `r`. */
  function rowState(d: number, r: number): RowState {
    const rect = sel.selRect.value
    const inRect = !!rect && r >= rect.top && r <= rect.bottom
    const a = s.selected.value
    const fl = nav.flash.value
    const now = deps.find.current.value
    const e = deps.editor.value
    return {
      height: geo.heightOf(d),
      sel: inRect ? { left: rect.left, right: rect.right, top: r === rect.top, below: Math.min(rect.bottom - r, mergeRows.value.get(r) ?? 1), multi: sel.multi.value } : null,
      active: a && a.row === r ? a.col : 0,
      hl: f.highlights.value.filter((h) => r >= h.ref.top && r <= h.bottom),
      hits: hitRows.value.get(r),
      hitNow: now && Number(now.slice(0, now.indexOf(':'))) === r ? Number(now.slice(now.indexOf(':') + 1)) : 0,
      flash: fl && fl.sheet === s.sheetName.value && r >= fl.rect.top && r <= fl.rect.bottom ? { left: fl.rect.left, right: fl.rect.right } : null,
      editor: e && e.row === r ? { col: e.col, bar: e.place === 'bar' } : null,
    }
  }

  /** Where a drawn column sits among its neighbors in the same pane, for overflow. */
  const colPlace = computed(() => {
    const map = new Map<number, { list: number[]; i: number }>()
    for (const list of [v.frozenColList.value, v.shownCols.value]) list.forEach((c, i) => map.set(c, { list, i }))
    return map
  })

  /** Whether the draft is empty: what a cell edited from the formula bar draws depends on that, not on its text. */
  const draftBlank = computed(() => !deps.editor.value?.text)

  /** A cell's style by its display row. */
  const styleOf = (d: number, c: number): CellStyle | undefined => w.styleAt(d, c, cellAddress(s.rowOf(d), c))
  const freezeRows = () => s.sheet.value?.freeze?.rows ?? 0
  const freezeCols = () => s.sheet.value?.freeze?.cols ?? 0

  /** A side's line: the cell's own, otherwise the neighbor's facing side (one line in Excel); otherwise a fill hides the gridline. */
  function edge(own: BorderLine | undefined, other: BorderLine | undefined, fill: string | undefined): { css?: CssLine; width: number } {
    const line = own ?? other
    if (line) return { css: borderCss(line), width: borderWidth(line) }
    return { css: fill ? { width: '1px', style: 'solid', color: fill } : undefined, width: 1 }
  }

  /** Whether a neighbor may take this cell's overflowing text: loaded, empty, not merged, not being edited. */
  function free(d: number, c: number, row: RowState): boolean {
    const r = s.rowOf(d)
    if (!w.emptyAt(d, c) || spanOf(r, c) !== undefined || ed.pendingAt(r, c) !== undefined) return false
    return row.editor?.col !== c
  }
  /** Text overflow as Excel draws it: left-aligned text runs right over empty cells, right-aligned text left, centered text both ways. */
  function overflowOf(d: number, c: number, h: Horizontal, bg: string, row: RowState): CellView['ov'] {
    const place = colPlace.value.get(c)
    if (!place) return undefined
    const run = (dir: 1 | -1, need?: (col: number) => boolean) => {
      let px = 0
      for (let i = place.i + dir; i >= 0 && i < place.list.length && px < OVERFLOW_MAX; i += dir) {
        const col = place.list[i]!
        if (!free(d, col, row) || (need && !need(col))) break
        px += geo.widthOf(col)
      }
      return px
    }
    if (h === 'centerContinuous') {
      const right = run(1, (col) => styleOf(d, col)?.alignment?.horizontal === 'centerContinuous')
      return right ? { left: 0, right, justify: 'center', bg } : undefined
    }
    if (h === 'left') {
      const right = run(1)
      return right ? { left: 0, right, justify: 'flex-start', bg } : undefined
    }
    if (h === 'right') {
      const left = run(-1)
      return left ? { left, right: 0, justify: 'flex-end', bg } : undefined
    }
    if (h === 'center') {
      const side = Math.min(run(1), run(-1))
      return side ? { left: side, right: side, justify: 'center', bg } : undefined
    }
    return undefined
  }

  /** Background layers, top first: the note indicator, selection and highlight tints, the data bar at the bottom. */
  function layers(css: Record<string, string>, list: Layer[]): void {
    if (!list.length) return
    css.backgroundImage = list.map((l) => l.img).join(', ')
    css.backgroundSize = list.map((l) => l.size).join(', ')
    css.backgroundPosition = list.map((l) => l.pos).join(', ')
    css.backgroundRepeat = 'no-repeat'
  }

  /** Everything needed to draw cell (r, c); `d` is its display row, `frozen` whether its row is frozen, `row` its row's state (`rowState`). */
  function cellView(d: number, r: number, c: number, frozen: boolean, row: RowState): CellView {
    const addr = cellAddress(r, c)
    const at = cellRect(r, c)
    const data = w.cellData(d, c, addr)
    const span = spanOf(r, c)
    // The merge's last row: by display index (geometry, neighbors) and by sheet row (selection, tables); under a filter merges are one row.
    const d2 = span ? d + span.rs - 1 : d
    const r2 = span ? r + span.rs - 1 : r
    const c2 = span ? c + span.cs - 1 : c
    const st = data.style
    const m = s.messages.value
    const pendingText = ed.pendingAt(r, c)
    const look = data.look
    // A rule that hides the value leaves only the bar or the icon; the tooltip still has the value.
    const barDraft = pendingText === undefined && !!row.editor && row.editor.bar && row.editor.col === c
    const shownText = pendingText !== undefined ? (pendingText ?? '') : barDraft || look?.hideValue ? '' : data.text
    const hasText = barDraft ? !draftBlank.value : shownText !== ''

    const css: Record<string, string> = fontCss(st, data.color)
    if (frozen) css.top = v.stickyTop(d)
    if (c <= v.frozenCols.value) css.left = v.stickyLeft(c)

    // Borders: right and bottom are the cell's own gridlines; top and left only at the sheet's edge (otherwise the neighbor draws them).
    const rightStyle = span ? styleOf(d, c2) : st
    const bottomStyle = span ? styleOf(d2, c) : st
    const colRight = geo.nextCol(c2, 1)
    const rowBelow = geo.nextRow(d2, 1)
    const nextC = colRight !== c2 ? styleOf(d, colRight) : undefined
    const nextR = rowBelow !== d2 ? styleOf(rowBelow, c) : undefined
    const right = edge(rightStyle?.border?.right, nextC?.border?.left, c2 === freezeCols() ? undefined : (rightStyle?.fill?.color ?? nextC?.fill?.color))
    const bottom = edge(bottomStyle?.border?.bottom, nextR?.border?.top, d2 === freezeRows() ? undefined : (bottomStyle?.fill?.color ?? nextR?.fill?.color))
    if (right.css) setSide(css, 'Right', right.css)
    if (bottom.css) setSide(css, 'Bottom', bottom.css)
    let topWidth = 0
    if (st?.border?.top && !geo.hasRow(d, -1)) {
      setSide(css, 'Top', borderCss(st.border.top)!)
      topWidth = borderWidth(st.border.top)
    }
    if (st?.border?.left && !geo.hasCol(c, -1)) setSide(css, 'Left', borderCss(st.border.left)!)

    // Insets: formula ranges, the selection's border and table borders; tints and bars as background layers.
    const shadows: string[] = []
    const bg: Layer[] = []
    const full = '100% 100%'
    if (data.note) bg.push({ img: 'linear-gradient(to bottom left, var(--xs-note) 50%, transparent 50%)', size: '7px 7px', pos: 'right top' })
    const rs = row.sel
    const inSel = !!rs && c >= rs.left && c <= rs.right
    const active = row.active === c
    if (rs && inSel && rs.multi) {
      if (!active) bg.push({ img: 'linear-gradient(var(--xs-sel-tint), var(--xs-sel-tint))', size: full, pos: '0 0' })
      if (rs.top) shadows.push('inset 0 1px 0 0 var(--xs-accent)')
      if (r2 - r >= rs.below) shadows.push('inset 0 -1px 0 0 var(--xs-accent)')
      if (c === rs.left) shadows.push('inset 1px 0 0 0 var(--xs-accent)')
      if (c2 >= rs.right) shadows.push('inset -1px 0 0 0 var(--xs-accent)')
    }
    const hit = !!row.hits?.has(c)
    if (hit) bg.push({ img: 'linear-gradient(var(--xs-hit), var(--xs-hit))', size: full, pos: '0 0' })
    const hl = row.hl.length ? f.highlightAt(r, c, row.hl) : null
    if (hl) {
      shadows.push(...hl.edges)
      for (const t of hl.tints) bg.push({ img: t, size: full, pos: '0 0' })
    }
    let tableHead = false
    for (const t of w.tables.value) {
      if (!rectContains(t.rect, at)) continue
      if (r === t.headerRow) tableHead = true
      if (r === t.rect.top) shadows.push('inset 0 1px 0 0 var(--xs-table-line)')
      if (r2 >= t.rect.bottom) shadows.push('inset 0 -1px 0 0 var(--xs-table-line)')
      if (c === t.rect.left) shadows.push('inset 1px 0 0 0 var(--xs-table-line)')
      if (c2 >= t.rect.right) shadows.push('inset -1px 0 0 0 var(--xs-table-line)')
    }
    if (look?.bar) bg.push(...barLayers(look.bar))
    if (shadows.length) css.boxShadow = shadows.join(', ')
    layers(css, bg)

    const lock = ed.lockAt(r, c)
    // List mode: a drill-down cell, or a first-column cell below the header that opens its row.
    const rowLink = !data.drill && c === 1 && deps.openRows() && r > freezeRows()
    const cls = [
      'xs-cell',
      active ? 'xs-active' : '',
      inSel ? 'xs-sel' : '',
      row.flash && c >= row.flash.left && c <= row.flash.right ? 'xs-flash' : '',
      hit ? 'xs-hit' : '',
      hit && row.hitNow === c ? 'xs-hit-now' : '',
      pendingText !== undefined ? 'xs-pending' : '',
      lock !== null ? 'xs-locked' : '',
      tableHead ? 'xs-thead' : '',
      frozen && c <= v.frozenCols.value ? 'xs-fz-corner' : frozen ? 'xs-fz-row' : c <= v.frozenCols.value ? 'xs-fz-col' : '',
      // The workbook's freeze edge is always shown: with the lock on it separates the sticky panes, with it off it shows what the lock freezes.
      d2 === freezeRows() ? 'xs-fz-edge-b' : '',
      c2 === freezeCols() ? 'xs-fz-edge-r' : '',
      data.drill ? 'xs-drill' : '',
      rowLink ? 'xs-rowlink' : '',
    ]

    const titleParts = [data.formula !== undefined ? `=${data.formula}\n${data.text}` : data.text]
    if (data.note) titleParts.push(`${m.note(data.note.author)}: ${data.note.text}`)
    if (lock !== null) titleParts.push(m.locked(lock))
    if (pendingText !== undefined) titleParts.push(m.pending)

    // The content box: exactly the row's height (a <tr> height is a minimum), with text clipped as in Excel.
    const h = horizontalOf(st, data.value)
    const wraps = wrapsOf(st)
    const vAlign = verticalOf(st)
    const boxH = Math.max(0, (span ? geo.spanHeight(d, span.rs) : row.height) - CELL_PAD_V * 2 - bottom.width - topWidth)
    const box: CellView['box'] = {
      class: ['xs-in', wraps ? 'xs-wrap' : '', vAlign === 'top' ? 'xs-top' : vAlign === 'center' ? 'xs-mid' : 'xs-bot'],
      style: { ...alignCss(h, st?.alignment?.indent ?? 0), height: `${boxH}px` },
    }

    const view: CellView = {
      attrs: { class: cls, style: css, title: rowLink ? m.openRow : titleParts.filter(Boolean).join('\n\n'), 'data-cell': `${r}:${c}` },
      box,
      kind: 'text',
      text: shownText,
      deco: decorationCss(st, false),
    }
    if (barDraft) view.draft = {}
    if (span) {
      view.attrs.rowspan = span.rsDom
      view.attrs.colspan = span.csDom
    }
    if (look?.icon) {
      view.icon = look.icon
      box.class.push('xs-iconic')
      // The icon alone follows the cell's alignment (General at the left), as Excel's Show Icon Only does.
      if (!hasText) {
        view.kind = 'icon'
        const own = st?.alignment?.horizontal
        box.style.justifyContent = own === 'center' || own === 'centerContinuous' ? 'center' : own === 'right' ? 'flex-end' : 'flex-start'
        return view
      }
    }
    const link = data.link
    if (link && pendingText === undefined) {
      view.link = link
      view.deco = decorationCss(st, true)
      view.kind = link.location ? 'place' : safeHref(link.href) ? 'href' : (link.href || link.target) && deps.hostLinks() ? 'host' : 'dead'
      return view
    }
    if (h === 'fill') {
      view.kind = 'fill'
      if (view.draft) view.draft.fill = geo.spanWidth(c, span?.cs ?? 1)
      else view.text = fillText(shownText, geo.spanWidth(c, span?.cs ?? 1))
    } else if (!span && !wraps && hasText && (overflows(data.value) || h === 'centerContinuous')) {
      const ov = overflowOf(d, c, h, st?.fill?.color ?? 'var(--xs-bg)', row)
      if (ov) {
        view.kind = 'overflow'
        view.ov = ov
        cls.push('xs-spill')
      }
    }
    return view
  }

  return { spanOf, coveredByMerge, rowState, cellView }
}

export type CellViews = ReturnType<typeof useCellView>
