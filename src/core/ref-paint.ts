// A formula's references coloured on one run of text by the CSS Custom Highlight API, as the formula bar's layer needs: elements around them would shape the text apart from the field's under it.

/** A piece of a text node to colour: its offsets and colour. */
export interface RefPaint {
  from: number
  to: number
  color: string
}
/** Ranges painted, each with its highlight's name, to take back on the next paint. */
export type Painted = [name: string, range: Range][]

/** The highlight's name of each colour, shared by every bar on the page (the registry is the document's). */
const names = new Map<string, string>()
/** The rules of the colours, adopted by each document or shadow root a bar is in. */
const sheets = new WeakMap<Document | ShadowRoot, CSSStyleSheet>()
const allSheets: CSSStyleSheet[] = []
const rule = (color: string, name: string) => `.xs-fx-layer::highlight(${name}) { color: ${color}; }`

/** The browser paints highlights: else the layer colours references with spans. */
export const canPaint = (): boolean => typeof CSS !== 'undefined' && 'highlights' in CSS && typeof Highlight === 'function' && typeof CSSStyleSheet === 'function'

function nameOf(color: string, root: Document | ShadowRoot): string {
  let sheet = sheets.get(root)
  if (!sheet) {
    sheet = new CSSStyleSheet()
    for (const [c, n] of names) sheet.insertRule(rule(c, n))
    root.adoptedStyleSheets = [...root.adoptedStyleSheets, sheet]
    sheets.set(root, sheet)
    allSheets.push(sheet)
  }
  let name = names.get(color)
  if (!name) {
    name = `xs-ref-${names.size}`
    names.set(color, name)
    for (const each of allSheets) each.insertRule(rule(color, name))
  }
  return name
}

/** Colour `pieces` of `node` in place of what `before` painted; null takes the old paint away only. */
export function paintRefs(node: Text | null, pieces: RefPaint[], before: Painted): Painted {
  for (const [name, range] of before) CSS.highlights.get(name)?.delete(range)
  if (!node) return []
  const root = node.getRootNode() as Document | ShadowRoot
  return pieces.map(({ from, to, color }) => {
    const name = nameOf(color, root)
    const range = new Range()
    range.setStart(node, from)
    range.setEnd(node, to)
    let highlight = CSS.highlights.get(name)
    if (!highlight) CSS.highlights.set(name, (highlight = new Highlight()))
    highlight.add(range)
    return [name, range]
  })
}
