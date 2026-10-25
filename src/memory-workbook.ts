// memorySource's dialogs over the in-memory workbook beside cell values: clearing, hyperlinks, notes, defined names and spelling, without an engine.
import { cellAddress, parseCell, parseRange } from 'xlcalc/a1'
import type { CellUpdate, EditOutcome } from './contract/source'
import type { MemoryCell } from './contract/memory'
import { misspelt } from './spelling'
import type { ClearKind, SheetLink, SheetName, SheetNote, SpellChecker, SpellingHit, SpellingRequest } from './contract/workbook'

/** A sheet of the in-memory workbook as these methods change it. */
export interface MemoryCells {
  cells: Map<string, MemoryCell>
}

/** The rectangle's cells to change: every cell of a small one, as Excel does, and only the existing cells of a large one (entire columns). */
const MAX_FILLED = 100_000

function cellsIn(store: MemoryCells, range: string, all: boolean): string[] | null {
  const r = parseRange(range)
  if (!r) return null
  if (all && (r.bottom - r.top + 1) * (r.right - r.left + 1) <= MAX_FILLED) {
    const out: string[] = []
    for (let row = r.top; row <= r.bottom; row++) for (let col = r.left; col <= r.right; col++) out.push(cellAddress(row, col))
    return out
  }
  return [...store.cells.keys()].filter((addr) => {
    const at = parseCell(addr)
    return at && at.row >= r.top && at.row <= r.bottom && at.col >= r.left && at.col <= r.right
  })
}

const noSheet = (sheet: string): EditOutcome => ({ ok: false, reason: 'failed', message: `No sheet ${sheet}` })
const notRange = (range: string): EditOutcome => ({ ok: false, reason: 'failed', message: `Not a range: ${range}` })

/** A name as Excel accepts one: a letter, `_` or `\` first, then letters, digits, `_`, `.`, `\` and `?`, not a cell address, at most 255 characters. */
const NAME = /^[\p{L}_\\][\p{L}\p{N}_.\\?]*$/u
const nameRefusal = (n: SheetName): string | null =>
  !NAME.test(n.name) || n.name.length > 255 || parseCell(n.name) || /^(r|c|true|false)$/i.test(n.name) ? `The name ${n.name} is not valid` : !n.ref.trim() ? `The name ${n.name} refers to nothing` : null

/** The in-memory workbook's methods; `measure` updates a sheet's extent after cells are removed. */
export function memoryWorkbook(storeOf: (sheet: string) => MemoryCells | undefined, measure: (store: MemoryCells) => void, initialNames: SheetName[], checker?: SpellChecker) {
  let names = initialNames.map((n) => ({ ...n }))

  /** Clear Formats removes a cell's style, number format and color; Clear All removes the cell. */
  function clear(sheet: string, ranges: string[], what: ClearKind): EditOutcome {
    const store = storeOf(sheet)
    if (!store) return noSheet(sheet)
    const changes: CellUpdate[] = []
    for (const range of ranges) {
      const cells = cellsIn(store, range, false)
      if (!cells) return notRange(range)
      for (const addr of cells) {
        const cell = store.cells.get(addr)
        if (!cell) continue
        if (what === 'all') {
          store.cells.delete(addr)
          changes.push({ sheet, cell: addr, text: '', value: null, formula: null })
          continue
        }
        const { style: _s, numberFormat: _f, color: _c, appearance: _a, ...rest } = cell
        store.cells.set(addr, rest)
      }
    }
    measure(store)
    return { ok: true, changes }
  }

  function link(sheet: string, range: string, link: SheetLink | null): EditOutcome {
    const store = storeOf(sheet)
    if (!store) return noSheet(sheet)
    if (link && !link.href && !link.location) return { ok: false, reason: 'refused', message: 'A hyperlink needs an address or a place in the workbook' }
    const cells = cellsIn(store, range, link !== null)
    if (!cells) return notRange(range)
    for (const addr of cells) {
      const { link: _l, ...rest } = store.cells.get(addr) ?? {}
      if (link) store.cells.set(addr, { ...rest, link: { ...link } })
      else if (Object.keys(rest).length) store.cells.set(addr, rest)
      else store.cells.delete(addr)
    }
    measure(store)
    return { ok: true, changes: [] }
  }

  function note(sheet: string, cell: string, note: SheetNote | null): EditOutcome {
    const store = storeOf(sheet)
    const at = parseCell(cell)
    if (!store) return noSheet(sheet)
    if (!at) return { ok: false, reason: 'failed', message: `No cell ${sheet}!${cell}` }
    const addr = cellAddress(at.row, at.col)
    const { note: _n, ...rest } = store.cells.get(addr) ?? {}
    if (note) store.cells.set(addr, { ...rest, note: { ...note } })
    else if (Object.keys(rest).length) store.cells.set(addr, rest)
    else store.cells.delete(addr)
    measure(store)
    return { ok: true, changes: [] }
  }

  /** The names replaced as given, each checked as Excel checks a name; formulas are kept as text (there is no engine). */
  function setNames(list: SheetName[]): EditOutcome {
    const seen = new Set<string>()
    for (const n of list) {
      const refused = nameRefusal(n)
      const key = `${n.scope?.toLocaleLowerCase() ?? ''}!${n.name.toLocaleLowerCase()}`
      if (refused || seen.has(key)) return { ok: false, reason: 'refused', message: refused ?? `The name ${n.name} already exists` }
      seen.add(key)
    }
    names = list.map((n) => ({ ...n }))
    return { ok: true, changes: [] }
  }

  /** Text constants by rows from the active cell, which is checked from `from`, wrapping around to the cells before it. */
  function spelling(req: SpellingRequest): SpellingHit | null {
    const store = storeOf(req.sheet)
    if (!store || !checker) return null
    const area = req.range ? parseRange(req.range) : null
    const ignore = new Set((req.ignore ?? []).map((w) => w.toLocaleLowerCase()))
    const cells = [...store.cells.entries()]
      .map(([addr, cell]) => ({ addr, cell, at: parseCell(addr)! }))
      .filter((c) => !area || (c.at.row >= area.top && c.at.row <= area.bottom && c.at.col >= area.left && c.at.col <= area.right))
      .sort((a, b) => a.at.row - b.at.row || a.at.col - b.at.col)
    const start = req.after ? parseCell(req.after) : null
    const from = start ? cells.findIndex((c) => c.at.row > start.row || (c.at.row === start.row && c.at.col >= start.col)) : 0
    const order = from > 0 ? [...cells.slice(from), ...cells.slice(0, from)] : cells
    for (const { addr, cell, at } of order) {
      if (cell.formula !== undefined || typeof cell.value !== 'string') continue
      const hit = misspelt(cell.value, start && at.row === start.row && at.col === start.col ? (req.from ?? 0) : 0, checker, ignore)
      if (hit) return { cell: addr, ...hit, suggestions: checker.suggest(hit.word) }
    }
    return null
  }

  return { clear, link, note, names: () => names.map((n) => ({ ...n })), setNames, spelling, learnWord: (word: string) => checker?.add?.(word) }
}
