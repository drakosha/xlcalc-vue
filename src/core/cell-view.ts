// One drawn cell: its classes, CSS and tooltip, the content box, borders shared with neighbours, merges and text overflow.
import { computed, type ComputedRef, type Ref } from 'vue'
import { cellAddress, rectContains } from 'xlcalc/a1'
import type { BorderLine, CellRect, CellStyle, Icon } from 'xlcalc/contract'
import { alignCss, borderCss, borderWidth, decorationCss, fillText, fontCss, horizontalOf, overflows, setSide, verticalOf, wrapsOf, type CssLine, type Horizontal } from './cell-style'
import { barLayers, type Layer } from './conditional'
import type { EditorState } from './editing'
import type { Editing } from './editing'
import type { Formula } from './formula'
import type { Geometry } from './geometry'
import type { Navigation } from './navigation'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Viewport } from './viewport'
import type { CellData, Windows } from './windows'

/** Vertical padding of a cell; with its borders it sits inside the row's height. */
const CELL_PAD_V = 1
/** How far text may overflow into empty neighbours, in pixels. */
const OVERFLOW_MAX = 2000

/** A merge as drawn: its size in the book (box height) and in drawn rows and columns (hidden ones are not in the DOM). */
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
  /** Underline and strike, set on the element that carries the text. */
  deco: Record<string, string>
  /** Overflow into empty neighbours: pixels taken left and right, the text's place, its background (hides crossed grid lines, as in Excel). */
  ov?: { left: number; right: number; justify: string; bg: string }
  link?: CellData['link']
  /** A conditional format's icon, drawn at the cell's left before its text. */
  icon?: Icon
}

/** Links a browser may open by itself; anything else (`javascript:` and the like) goes to the host or is shown only. */
export const safeHref = (href: string | undefined): boolean => !!href && /^(https?:|mailto:|tel:|ftp:)/i.test(href.trim())

export interface CellViewDeps {
  editor: Ref<EditorState | null>
  /** Search hits by `row:col`, and the current one. */
  find: { hitSet: ComputedRef<Set<string>>; current: ComputedRef<string> }
  /** Whether the host listens to `openLink`. */
  hostLinks: () => boolean
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
  /** Merges of the drawn piece by sheet rows: the top-left cell's span, null for covered cells. A merge not wholly drawn (or across the freeze edge) is drawn as plain cells; under a filter only one-row merges stay. */
  const merges = computed(() => {
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
      // The top-left cell may be hidden: the merge then starts at its first shown cell.
      const d0 = geo.heightOf(d1) > 0 ? d1 : geo.nextRow(d1, 1)
      const c0 = geo.widthOf(m.left) > 0 ? m.left : geo.nextCol(m.left, 1)
      for (let r = m.top; r <= m.bottom; r++) for (let c = m.left; c <= m.right; c++) map.set(`${r}:${c}`, null)
      map.set(`${s.rowOf(d0)}:${c0}`, { rs: d2 - d0 + 1, cs: m.right - c0 + 1, rsDom, csDom })
    }
    return map
  })
  const spanOf = (r: number, c: number) => merges.value.get(`${r}:${c}`)
  const coveredByMerge = (r: number, c: number) => merges.value.get(`${r}:${c}`) === null

  /** Where a drawn column sits among its neighbours of the same pane, for overflow. */
  const colPlace = computed(() => {
    const map = new Map<number, { list: number[]; i: number }>()
    for (const list of [v.frozenColList.value, v.shownCols.value]) list.forEach((c, i) => map.set(c, { list, i }))
    return map
  })

  /** A cell's style by its display row. */
  const styleOf = (d: number, c: number): CellStyle | undefined => w.styleAt(d, c, cellAddress(s.rowOf(d), c))
  const freezeRows = () => s.sheet.value?.freeze?.rows ?? 0
  const freezeCols = () => s.sheet.value?.freeze?.cols ?? 0

  /** A side's line: the cell's own, else the neighbour's facing side (one line in Excel); else a fill hides the grid line. */
  function edge(own: BorderLine | undefined, other: BorderLine | undefined, fill: string | undefined): { css?: CssLine; width: number } {
    const line = own ?? other
    if (line) return { css: borderCss(line), width: borderWidth(line) }
    return { css: fill ? { width: '1px', style: 'solid', color: fill } : undefined, width: 1 }
  }

  /** May a neighbour take this cell's overflowing text: loaded, empty, not merged, not being edited. */
  function free(d: number, c: number): boolean {
    const r = s.rowOf(d)
    if (!w.emptyAt(d, c) || spanOf(r, c) !== undefined || ed.pendingAt(r, c) !== undefined) return false
    const e = deps.editor.value
    return !(e && e.row === r && e.col === c)
  }
  /** Text overflow as Excel draws it: left-aligned text runs right over empty cells, right-aligned left, centred both ways. */
  function overflowOf(d: number, c: number, h: Horizontal, bg: string): CellView['ov'] {
    const place = colPlace.value.get(c)
    if (!place) return undefined
    const run = (dir: 1 | -1, need?: (col: number) => boolean) => {
      let px = 0
      for (let i = place.i + dir; i >= 0 && i < place.list.length && px < OVERFLOW_MAX; i += dir) {
        const col = place.list[i]!
        if (!free(d, col) || (need && !need(col))) break
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

  /** Background layers, top first: the note's corner, selection and highlight tints, the data bar at the bottom. */
  function layers(css: Record<string, string>, list: Layer[]): void {
    if (!list.length) return
    css.backgroundImage = list.map((l) => l.img).join(', ')
    css.backgroundSize = list.map((l) => l.size).join(', ')
    css.backgroundPosition = list.map((l) => l.pos).join(', ')
    css.backgroundRepeat = 'no-repeat'
  }

  /** Everything to draw cell (r, c); `d` is its display row, `frozen` whether its row is a frozen one. */
  function cellView(d: number, r: number, c: number, frozen: boolean): CellView {
    const addr = cellAddress(r, c)
    const at: CellRect = { top: r, left: c, bottom: r, right: c }
    const data = w.cellData(d, c, addr)
    const span = spanOf(r, c)
    // The merge's last row: display (geometry, neighbours) and sheet (selection, tables); under a filter merges are one row.
    const d2 = span ? d + span.rs - 1 : d
    const r2 = span ? r + span.rs - 1 : r
    const c2 = span ? c + span.cs - 1 : c
    const st = data.style
    const m = s.messages.value
    const pendingText = ed.pendingAt(r, c)
    const editing = deps.editor.value
    const look = data.look
    // A rule that hides the value leaves the bar or the icon alone; the tooltip still has the value.
    const shownText =
      pendingText !== undefined ? (pendingText ?? '') : editing && editing.place === 'bar' && editing.row === r && editing.col === c ? editing.text : look?.hideValue ? '' : data.text

    const css: Record<string, string> = fontCss(st, data.color)
    if (frozen) css.top = v.stickyTop(d)
    if (c <= v.frozenCols.value) css.left = v.stickyLeft(c)

    // Borders: right and bottom are the cell's own grid lines; top and left only at the sheet's edge (else the neighbour draws them).
    const rightStyle = span ? styleOf(d, c2) : st
    const bottomStyle = span ? styleOf(d2, c) : st
    const hasRight = geo.hasCol(c2, 1)
    const hasBelow = geo.hasRow(d2, 1)
    const nextC = hasRight ? styleOf(d, geo.nextCol(c2, 1)) : undefined
    const nextR = hasBelow ? styleOf(geo.nextRow(d2, 1), c) : undefined
    const right = edge(rightStyle?.border?.right, nextC?.border?.left, c2 === freezeCols() ? undefined : (rightStyle?.fill?.color ?? nextC?.fill?.color))
    const bottom = edge(bottomStyle?.border?.bottom, nextR?.border?.top, d2 === freezeRows() ? undefined : (bottomStyle?.fill?.color ?? nextR?.fill?.color))
    if (right.css) setSide(css, 'Right', right.css)
    if (bottom.css) setSide(css, 'Bottom', bottom.css)
    let topWidth = 0
    if (!geo.hasRow(d, -1) && st?.border?.top) {
      setSide(css, 'Top', borderCss(st.border.top)!)
      topWidth = borderWidth(st.border.top)
    }
    if (!geo.hasCol(c, -1) && st?.border?.left) setSide(css, 'Left', borderCss(st.border.left)!)

    // Insets: formula ranges, the selection's frame and tables' frames; tints and bars as background layers.
    const shadows: string[] = []
    const bg: Layer[] = []
    const full = '100% 100%'
    if (data.note) bg.push({ img: 'linear-gradient(to bottom left, var(--xs-note) 50%, transparent 50%)', size: '7px 7px', pos: 'right top' })
    const rect = sel.selRect.value
    const active = sel.isActive(r, c)
    if (rect && sel.multi.value && rectContains(rect, at)) {
      if (!active) bg.push({ img: 'linear-gradient(var(--xs-sel-tint), var(--xs-sel-tint))', size: full, pos: '0 0' })
      if (r === rect.top) shadows.push('inset 0 1px 0 0 var(--xs-accent)')
      if (r2 >= rect.bottom) shadows.push('inset 0 -1px 0 0 var(--xs-accent)')
      if (c === rect.left) shadows.push('inset 1px 0 0 0 var(--xs-accent)')
      if (c2 >= rect.right) shadows.push('inset -1px 0 0 0 var(--xs-accent)')
    }
    const hitKey = `${r}:${c}`
    if (deps.find.hitSet.value.has(hitKey)) bg.push({ img: 'linear-gradient(var(--xs-hit), var(--xs-hit))', size: full, pos: '0 0' })
    const hl = f.highlightAt(r, c)
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

    const fl = nav.flash.value
    const lock = ed.lockAt(r, c)
    const cls = [
      'xs-cell',
      active ? 'xs-active' : '',
      rect && rectContains(rect, at) ? 'xs-sel' : '',
      fl && rectContains(fl, at) ? 'xs-flash' : '',
      deps.find.hitSet.value.has(hitKey) ? 'xs-hit' : '',
      deps.find.current.value === hitKey ? 'xs-hit-now' : '',
      pendingText !== undefined ? 'xs-pending' : '',
      lock !== null ? 'xs-locked' : '',
      tableHead ? 'xs-thead' : '',
      frozen && c <= v.frozenCols.value ? 'xs-fz-corner' : frozen ? 'xs-fz-row' : c <= v.frozenCols.value ? 'xs-fz-col' : '',
      // The book's freeze edge shows always: with the lock it parts the sticky panes, without it it shows what the lock freezes.
      d2 === freezeRows() ? 'xs-fz-edge-b' : '',
      c2 === freezeCols() ? 'xs-fz-edge-r' : '',
    ]

    const titleParts = [data.formula !== undefined ? `=${data.formula}\n${data.text}` : data.text]
    if (data.note) titleParts.push(`${m.note(data.note.author)}: ${data.note.text}`)
    if (lock !== null) titleParts.push(m.locked(lock))
    if (pendingText !== undefined) titleParts.push(m.pending)

    // The content box: exactly the row's height (a <tr> height is a minimum), text clipped as in Excel.
    const h = horizontalOf(st, data.value)
    const wraps = wrapsOf(st)
    const vAlign = verticalOf(st)
    const boxH = Math.max(0, geo.spanHeight(d, span?.rs ?? 1) - CELL_PAD_V * 2 - bottom.width - topWidth)
    const box: CellView['box'] = {
      class: ['xs-in', wraps ? 'xs-wrap' : '', vAlign === 'top' ? 'xs-top' : vAlign === 'center' ? 'xs-mid' : 'xs-bot'],
      style: { ...alignCss(h, st?.alignment?.indent ?? 0), height: `${boxH}px` },
    }

    const view: CellView = {
      attrs: { class: cls, style: css, title: titleParts.filter(Boolean).join('\n\n'), 'data-cell': `${r}:${c}` },
      box,
      kind: 'text',
      text: shownText,
      deco: decorationCss(st, false),
    }
    if (span) {
      view.attrs.rowspan = span.rsDom
      view.attrs.colspan = span.csDom
    }
    if (look?.icon) {
      view.icon = look.icon
      box.class.push('xs-iconic')
      // The icon alone stands by the cell's alignment (General at the left), as Excel's Show Icon Only.
      if (!shownText) {
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
      view.text = fillText(shownText, geo.spanWidth(c, span?.cs ?? 1))
    } else if (!span && !wraps && shownText && (overflows(data.value) || h === 'centerContinuous')) {
      const ov = overflowOf(d, c, h, st?.fill?.color ?? 'var(--xs-bg)')
      if (ov) {
        view.kind = 'overflow'
        view.ov = ov
        cls.push('xs-spill')
      }
    }
    return view
  }

  return { merges, spanOf, coveredByMerge, cellView }
}

export type CellViews = ReturnType<typeof useCellView>
