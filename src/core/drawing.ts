// Drawings over the cells (slicers, timelines): a place as xlcalc gives it, the cells of two corners with offsets in EMU,
// to the grid's pixels and back, and the parts of a box each pane of frozen panes draws, as Excel draws an object in each pane.
import { cellAddress, parseCell } from 'xlcalc/a1'
import type { DrawingPlace } from 'xlcalc/contract'
import type { Geometry } from './geometry'
import type { Viewport } from './viewport'

/** EMU in a pixel at 96 dpi (914,400 to an inch). */
export const EMU_PX = 9525

/** A box in the offsets' space of the grid: pixels from the first row's top and the first column's left, unzoomed. */
export interface DrawnBox {
  x: number
  y: number
  width: number
  height: number
}

interface RowMap {
  rowOf: (d: number) => number
  indexOf: (r: number) => number | undefined
  /** The display filter shows rows in another order than the sheet's: a box takes its height from its sheet rows. */
  filtered?: boolean
}

/** The left edge of a column, also of one beyond the columns drawn (by their widths, the default for most). */
function colEdge(geo: Geometry, col: number): number {
  const offsets = geo.offsets.value
  const count = offsets.length - 1
  if (col - 1 <= count) return offsets[col - 1] ?? 0
  let x = offsets[count] ?? 0
  for (let c = count + 1; c < col; c++) x += geo.widthOf(c)
  return x
}
/** The top edge of a display row, also of one beyond the rows drawn. */
function rowEdge(geo: Geometry, d: number): number {
  const offsets = geo.rowOffsets.value
  const count = offsets.length - 1
  if (d - 1 <= count) return offsets[d - 1] ?? 0
  let y = offsets[count] ?? 0
  for (let r = count + 1; r < d; r++) y += geo.heightOf(r)
  return y
}
/** An offset into a cell in EMU as pixels, no more than the cell's size: an anchor's offset lies within its cell (ECMA-376's `colOff`, `rowOff`). */
const within = (emu: number | undefined, size: number) => Math.min(Math.max(0, (emu ?? 0) / EMU_PX), size)

/**
 * Where a place's corners are drawn; null for a place that is not one, or whose top-left row the display filter leaves out.
 * Under the display filter the box keeps the height of its sheet rows from its top-left row.
 */
export function boxOfPlace(place: DrawingPlace, geo: Geometry, rows: RowMap): DrawnBox | null {
  const from = parseCell(place.from)
  const to = parseCell(place.to)
  if (!from || !to) return null
  const d1 = rows.indexOf(from.row)
  if (d1 === undefined) return null
  const x = colEdge(geo, from.col) + within(place.fromOffset?.[0], geo.widthOf(from.col))
  const y = rowEdge(geo, d1) + within(place.fromOffset?.[1], geo.heightOf(d1))
  const x2 = colEdge(geo, to.col) + within(place.toOffset?.[0], geo.widthOf(to.col))
  let y2: number
  if (rows.filtered) {
    // the sheet rows from its top-left row down to its bottom-right one, those the filter leaves out as hidden rows
    const heightOf = (r: number) => {
      const d = rows.indexOf(r)
      return d === undefined ? 0 : geo.heightOf(d)
    }
    let h = 0
    for (let r = from.row; r < to.row; r++) h += heightOf(r)
    y2 = rowEdge(geo, d1) + h + within(place.toOffset?.[1], heightOf(to.row))
  } else y2 = rowEdge(geo, to.row) + within(place.toOffset?.[1], geo.heightOf(to.row))
  return { x, y, width: Math.max(0, x2 - x), height: Math.max(0, y2 - y) }
}

/** The cell under a point of the offsets' space with the point's offset into it in EMU. */
function anchorAt(x: number, y: number, geo: Geometry, rows: RowMap): { cell: string; offset: [number, number] } {
  const px = Math.max(0, x)
  const py = Math.max(0, y)
  const col = geo.colAt(px)
  const d = geo.rowAt(py)
  const dx = Math.max(0, px - (geo.offsets.value[col - 1] ?? 0))
  const dy = Math.max(0, py - (geo.rowOffsets.value[d - 1] ?? 0))
  return { cell: cellAddress(rows.rowOf(d), col), offset: [Math.round(dx * EMU_PX), Math.round(dy * EMU_PX)] }
}

/** A box of the offsets' space as a place: its corners' cells and offsets, as Excel anchors a slicer moved or resized. */
export function placeOfBox(box: DrawnBox, geo: Geometry, rows: RowMap): DrawingPlace {
  const a = anchorAt(box.x, box.y, geo, rows)
  const b = anchorAt(box.x + box.width, box.y + box.height, geo, rows)
  return { from: a.cell, to: b.cell, fromOffset: a.offset, toOffset: b.offset }
}

/**
 * A part of a box a pane draws: where it is in the scroller's content, the inset that clips it to the pane, and its layer
 * (`frozen` parts stick and lie over the frozen cells; the body's part scrolls under them and the headings).
 */
export interface DrawnPart {
  key: string
  top: number
  left: number
  clip: string
  z: number
}

/** The layers of the parts: the body's under every frozen cell; a frozen row's or column's over its cells; the corner's over both. */
const LAYERS = { body: 1, cols: 2, rows: 3, corner: 5 }

/** The parts of a box in the panes of frozen panes (one without them): each pane clips its part and scrolls as its cells do. */
export function partsOf(box: DrawnBox, v: Viewport): DrawnPart[] {
  const fh = v.frozenH.value
  const fw = v.frozenW.value
  const scrollTop = v.viewBox.value.top - v.headH.value
  const scrollLeft = v.viewBox.value.left - v.leadW.value
  const out: DrawnPart[] = []
  for (const rowsFrozen of fh > 0 ? [true, false] : [false]) {
    for (const colsFrozen of fw > 0 ? [true, false] : [false]) {
      const top = rowsFrozen ? 0 : fh
      const bottom = rowsFrozen ? fh : Infinity
      const left = colsFrozen ? 0 : fw
      const right = colsFrozen ? fw : Infinity
      if (box.y >= bottom || box.y + box.height <= top || box.x >= right || box.x + box.width <= left) continue
      const inset = [Math.max(0, top - box.y), Math.max(0, box.x + box.width - right), Math.max(0, box.y + box.height - bottom), Math.max(0, left - box.x)]
      out.push({
        key: `${rowsFrozen ? 'r' : ''}${colsFrozen ? 'c' : ''}`,
        top: v.headH.value + box.y + (rowsFrozen ? scrollTop : 0),
        left: v.leadW.value + box.x + (colsFrozen ? scrollLeft : 0),
        clip: inset.some((n) => n > 0) ? `inset(${inset.map((n) => `${n}px`).join(' ')})` : '',
        z: rowsFrozen && colsFrozen ? LAYERS.corner : rowsFrozen ? LAYERS.rows : colsFrozen ? LAYERS.cols : LAYERS.body,
      })
    }
  }
  return out
}
