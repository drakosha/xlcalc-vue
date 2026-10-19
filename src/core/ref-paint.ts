// A formula's references colored on a single text run by the CSS Custom Highlight API, as the formula bar's layer needs (elements around them would shape the text differently from the input's text under it), and which one lies under a point.

/** A piece of a text node to color: its offsets and color, underlined while a held Ctrl or Cmd would navigate to it. */
export interface RefPaint {
  from: number
  to: number
  color: string
  underline?: boolean
}
/** Painted ranges, each with its highlight's name, to remove on the next paint. */
export type Painted = [name: string, range: Range][]

/** The highlight name for each color (underlined ones separate), shared by every formula bar on the page (the registry belongs to the document). */
const names = new Map<string, string>()
/** The color rules, adopted by each document or shadow root a formula bar is in. */
const sheets = new WeakMap<Document | ShadowRoot, CSSStyleSheet>()
/** The same style sheets for a new color's rule, held weakly: a removed shadow root takes its sheet with it. */
const allSheets = new Set<WeakRef<CSSStyleSheet>>()
const rule = (key: string, name: string) => {
  const [color, underline] = key.split(' ')
  return `.xs-fx-layer::highlight(${name}) { color: ${color};${underline ? ' text-decoration: underline;' : ''} }`
}

/** Whether the browser paints highlights; otherwise the layer colors references with spans. */
export const canPaint = (): boolean => typeof CSS !== 'undefined' && 'highlights' in CSS && typeof Highlight === 'function' && typeof CSSStyleSheet === 'function'

function nameOf(key: string, root: Document | ShadowRoot): string {
  let sheet = sheets.get(root)
  if (!sheet) {
    sheet = new CSSStyleSheet()
    for (const [c, n] of names) sheet.insertRule(rule(c, n))
    root.adoptedStyleSheets = [...root.adoptedStyleSheets, sheet]
    sheets.set(root, sheet)
    allSheets.add(new WeakRef(sheet))
  }
  let name = names.get(key)
  if (!name) {
    name = `xs-ref-${names.size}`
    names.set(key, name)
    for (const held of allSheets) {
      const each = held.deref()
      if (each) each.insertRule(rule(key, name))
      else allSheets.delete(held)
    }
  }
  return name
}

/** Colors `pieces` of `node` in place of what `before` painted; null only removes the old paint. */
export function paintRefs(node: Text | null, pieces: RefPaint[], before: Painted): Painted {
  for (const [name, range] of before) CSS.highlights.get(name)?.delete(range)
  if (!node) return []
  const root = node.getRootNode() as Document | ShadowRoot
  return pieces.map(({ from, to, color, underline }) => {
    const name = nameOf(underline ? `${color} u` : color, root)
    const range = new Range()
    range.setStart(node, from)
    range.setEnd(node, to)
    let highlight = CSS.highlights.get(name)
    if (!highlight) CSS.highlights.set(name, (highlight = new Highlight()))
    highlight.add(range)
    return [name, range]
  })
}

/** A range over `layer`'s text from one offset to another, walking its text nodes (spans or a single node). */
function rangeIn(layer: HTMLElement, from: number, to: number): Range {
  const range = new Range()
  const walker = document.createTreeWalker(layer, NodeFilter.SHOW_TEXT)
  let at = 0
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const end = at + node.textContent!.length
    if (from >= at && from <= end) range.setStart(node, from - at)
    if (to >= at && to <= end) return range.setEnd(node, to - at), range
    at = end
  }
  return range
}

/**
 * The piece of `layer`'s text under a point (of the mouse or a finger), by the boxes its characters are drawn in, its line's
 * height around them and `slop` pixels to either side; -1 when off all of them. Beyond a line's end or between lines there is none.
 */
export function pieceAtPoint(layer: HTMLElement, pieces: { from: number; to: number }[], x: number, y: number, slop = 0): number {
  const line = parseFloat(getComputedStyle(layer).lineHeight)
  return pieces.findIndex(({ from, to }) =>
    [...rangeIn(layer, from, to).getClientRects()].some((r) => {
      const half = Math.max(r.height, line || 0) / 2
      return x >= r.left - slop && x <= r.right + slop && Math.abs(y - (r.top + r.height / 2)) <= half
    }),
  )
}
