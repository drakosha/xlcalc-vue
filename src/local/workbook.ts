// The local source's dialogs over the workbook beside cell values: clearing, hyperlinks, notes, defined names, rules, tables,
// dependents and spelling, each a session call with its formulas translated between the stored form and the workbook's locale.
import { toLocal, validateLocalFormula, validateNames, type AsyncSession, type BookLayout, type NameInput, type CellChange, type ConditionalFormat, type DataValidation, type EditResult } from 'xlcalc'
import { cellAddress, parseCell } from 'xlcalc/a1'
import type { EditOutcome } from '../contract/source'
import type { ClearKind, SheetDependents, SheetLink, SheetName, SheetNote, SheetRules, SpellChecker, SpellingHit, SpellingRequest, TableOp } from '../contract/workbook'
import { failed } from './outcome'
import { misspelt } from '../spelling'

/** How the source completes a change: the windows re-read in place over ranges, or the sheets and windows read again. */
export interface Completion {
  repainted(got: EditResult, ranges: { sheet: string; range: string }[]): EditOutcome
  reloaded(changes: CellChange[]): EditOutcome
}

/** The thresholds of color scales, data bars and icon sets, whose `value` is a formula. */
const THRESHOLDS = new Set(['min', 'max', 'num', 'percent', 'percentile', 'formula', 'autoMin', 'autoMax'])

/** A rule with each of its formulas passed through `f`: `formula`, `formulas`, `formula1`, `formula2` and a threshold's `value`, at any depth. */
function mapFormulas<T>(rule: T, f: (formula: string) => string): T {
  const walk = (x: unknown): unknown => {
    if (Array.isArray(x)) return x.map(walk)
    if (!x || typeof x !== 'object') return x
    const o = x as Record<string, unknown>
    const out: Record<string, unknown> = {}
    const threshold = typeof o.type === 'string' && THRESHOLDS.has(o.type)
    for (const [k, v] of Object.entries(o)) {
      if (typeof v === 'string' && (k === 'formula' || k === 'formula1' || k === 'formula2' || (threshold && k === 'value'))) out[k] = f(v)
      else if (k === 'formulas' && Array.isArray(v)) out[k] = v.map((s) => (typeof s === 'string' ? f(s) : s))
      else out[k] = walk(v)
    }
    return out
  }
  return walk(rule) as T
}

/** A workbook's defined names as stored; `hidden` (xlsx's hidden names) is read by xlcalc 0.57 and later. */
const storedNames = (layout: BookLayout): (NameInput & { hidden?: boolean })[] => layout.names ?? []

/** A typed formula that does not parse, which refuses the change it is part of. */
class Unparsed extends Error {}

/** `run`'s outcome, or a refusal with the message of a formula in it that does not parse. */
async function parsed(run: () => Promise<EditOutcome>): Promise<EditOutcome> {
  try {
    return await run()
  } catch (e) {
    if (e instanceof Unparsed) return { ok: false, reason: 'refused', message: e.message }
    throw e
  }
}

/** The workbook methods of the local source over its session, completed by `done`, with formulas in `locale`. */
export function workbookMethods(locale: string, done: Completion, checker?: SpellChecker) {
  const local = (formula: string): string => {
    const t = toLocal(formula, locale)
    return t.ok ? t.formula : formula
  }
  /** Typed formula text as stored, checked as Excel checks a formula; a formula that does not parse rejects the change. */
  const stored = (text: string): string => {
    const t = validateLocalFormula(text, locale)
    if (!t.ok) throw new Unparsed(t.message)
    return t.canonical
  }

  /** Clear Formats and Clear All: the session's `clearCells`, which changes displayed text, so the sheet is read again. */
  async function clear(s: AsyncSession, sheet: string, ranges: string[], what: ClearKind): Promise<EditOutcome> {
    if (!ranges.length) return { ok: true, changes: [] }
    return done.reloaded((await s.clearCells(ranges.map((range) => ({ sheet, range })), what)).changes)
  }

  async function link(s: AsyncSession, sheet: string, range: string, link: SheetLink | null): Promise<EditOutcome> {
    if (link && !link.href && !link.location) return failed('A hyperlink needs an address or a place in the workbook')
    const set = link ? { url: link.href, location: link.location, tooltip: link.tooltip } : null
    const clean = set && Object.fromEntries(Object.entries(set).filter(([, v]) => v !== undefined))
    return done.repainted(await s.patchCells([{ sheet, range, set: { link: clean } }]), [{ sheet, range }])
  }

  async function note(s: AsyncSession, sheet: string, cell: string, note: SheetNote | null): Promise<EditOutcome> {
    const at = parseCell(cell)
    if (!at) return failed(`No cell ${sheet}!${cell}`)
    const addr = cellAddress(at.row, at.col)
    return done.repainted(await s.patchCells([{ sheet, cell: addr, set: { note } }]), [{ sheet, range: addr }])
  }

  /** The names the Name Manager lists: a hidden name (xlsx `hidden`, which xlcalc 0.57 reads) is not shown, as in Excel. */
  async function names(s: AsyncSession): Promise<SheetName[]> {
    return storedNames(await s.getLayout())
      .filter((n) => !n.hidden)
      .map((n) => ({ name: n.name, ref: local(n.ref), ...(n.scope ? { scope: n.scope } : {}) }))
  }

  /**
   * The names as the Name Manager leaves them, their formulas parsed in the workbook's locale and the list checked by the
   * library's `validateNames` (a name Excel rejects, a duplicate, an unknown scope), since `setNames` itself does not check
   * names yet; the hidden names, which the list does not show, are kept as they are. Names change values anywhere, so the
   * sheet is read again.
   */
  async function setNames(s: AsyncSession, list: SheetName[]): Promise<EditOutcome> {
    return parsed(async () => {
      const layout = await s.getLayout()
      const hidden = storedNames(layout).filter((n) => n.hidden)
      const input = [...list.map((n) => ({ name: n.name, ref: stored(n.ref), ...(n.scope ? { scope: n.scope } : {}) })), ...hidden]
      const tables = layout.sheets.flatMap((sh) => (sh.tables ?? []).map((t) => ({ name: t.name, columns: t.columns.map((c) => c.name) })))
      const dropped = validateNames(input, { sheets: layout.sheets.map((sh) => sh.name), tables, locale }).find((i) => i.dropped)
      if (dropped) return { ok: false, reason: 'refused', message: dropped.reason }
      return done.reloaded((await s.setNames(input)).changes)
    })
  }

  async function dependents(s: AsyncSession, sheet: string, range: string, signal?: AbortSignal): Promise<SheetDependents> {
    const got = await s.getDependents([{ sheet, range }], { maxReferences: 1, ...(signal ? { signal } : {}) })
    return { cells: got.dependents.map((d) => ({ sheet: d.sheet, cell: d.cell })), more: got.truncated }
  }

  async function rules(s: AsyncSession, sheet: string): Promise<SheetRules> {
    const got = (await s.getLayout()).sheets.find((x) => x.name.toLocaleLowerCase() === sheet.toLocaleLowerCase())
    if (!got) throw new Error(`No sheet ${sheet}`)
    return { conditionalFormats: (got.conditionalFormats ?? []).map((r) => mapFormulas(r, local)), validations: (got.validations ?? []).map((r) => mapFormulas(r, local)) }
  }

  /** The sheet's rules replaced in one `setLayout` call; conditional formats change how cells look, so the sheet is read again. */
  async function setRules(s: AsyncSession, sheet: string, next: Partial<SheetRules>): Promise<EditOutcome> {
    return parsed(async () => {
      const edits = [
        ...(next.conditionalFormats ? [{ sheet, conditionalFormats: next.conditionalFormats.map((r): ConditionalFormat => mapFormulas(r, stored)) }] : []),
        ...(next.validations ? [{ sheet, validations: next.validations.map((r): DataValidation => mapFormulas(r, stored)) }] : []),
      ]
      if (!edits.length) return { ok: true, changes: [] }
      return done.reloaded((await s.setLayout(edits)).changes)
    })
  }

  async function table(s: AsyncSession, op: TableOp): Promise<EditOutcome> {
    const got = await s.editStructure(op)
    return got.ok ? done.reloaded(got.changes) : { ok: false, reason: 'refused', code: got.code, message: got.message }
  }

  /** Excel's Spelling: the active cell's text from `from`, then the other text constants in Find's order by rows, wrapping around; formulas are skipped. */
  async function spelling(s: AsyncSession, req: SpellingRequest, signal?: AbortSignal): Promise<SpellingHit | null> {
    if (!checker) return null
    const ignore = new Set((req.ignore ?? []).map((w) => w.toLocaleLowerCase()))
    const hitIn = (c: { cell: string; value?: unknown; formula?: string | null }, from: number): SpellingHit | null => {
      if (c.formula || typeof c.value !== 'string') return null
      const hit = misspelt(c.value, from, checker, ignore)
      return hit && { cell: c.cell, ...hit, suggestions: checker.suggest(hit.word) }
    }
    const opts = { fields: ['value', 'formula'] as const, ...(signal ? { signal } : {}) }
    const after = req.after && parseCell(req.after) ? cellAddress(parseCell(req.after)!.row, parseCell(req.after)!.col) : undefined
    if (after) {
      const own = (await s.readCells([{ sheet: req.sheet, range: after }], { ...opts, fields: [...opts.fields] })).cells[0]
      const hit = own && hitIn(own, req.from ?? 0)
      if (hit) return hit
    }
    const used = await s.getUsedRange(req.sheet)
    const range = req.range ?? `A1:${cellAddress(Math.max(1, used.rows), Math.max(1, used.columns))}`
    let cursor: string | undefined
    do {
      const page = await s.findCells([{ sheet: req.sheet, range }], '*', { ...opts, fields: [...opts.fields], in: 'value', wildcards: true, ...(after ? { after: { sheet: req.sheet, cell: after } } : {}), ...(cursor ? { cursor } : {}) })
      for (const c of page.cells) {
        const hit = c.cell === after ? null : hitIn(c, 0)
        if (hit) return hit
      }
      cursor = page.cursor
    } while (cursor)
    return null
  }

  async function learnWord(word: string): Promise<void> {
    checker?.add?.(word)
  }

  return { clear, link, note, names, setNames, dependents, rules, setRules, table, spelling, learnWord }
}
