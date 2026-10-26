// Formulas of conditional formatting rules in the workbook's locale: shifting relative references between the active cell,
// which Excel's dialogs show them for, and the rule's first cell, which they are written for; what a value box's text
// stores; and the formula Excel writes for a text rule whose text is a reference.
import { colLetters, colNumber, MAX_COLS, MAX_ROWS } from 'xlcalc/a1'
import type { ConditionalFormat, Locale, ScanOptions } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'
import type { Entry } from '../entry'
import { isGridRef } from '../grid-ref'

/** One end of a reference: a cell (`$A1`), a column (`$A`) or a row (`1`), with `$` where it is absolute. */
const END = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/

/** A row or column moved by `by`, coming round the other edge past the sheet's, as a rule's formula computes it. */
const wrap = (n: number, by: number, max: number) => ((((n - 1 + by) % max) + max) % max) + 1

function shiftEnd(end: string, rows: number, cols: number): string {
  const m = END.exec(end)
  if (!m || (!m[2] && !m[4])) return end
  const [, colAbs, col, rowAbs, row] = m
  let out = ''
  if (col) out += colAbs ? `$${col}` : colLetters(wrap(colNumber(col)!, cols, MAX_COLS))
  if (row) out += rowAbs ? `$${row}` : String(wrap(Number(row), rows, MAX_ROWS))
  return out
}

/** A reference's text with its relative parts moved: `A1:$B2` by one row is `A2:$B3`; a sheet name and a spill `#` are kept. */
function shiftRef(text: string, rows: number, cols: number): string {
  const bang = text.lastIndexOf('!')
  const sheet = text.slice(0, bang + 1)
  const spill = text.endsWith('#') ? '#' : ''
  const body = text.slice(bang + 1, text.length - spill.length)
  return sheet + body.split(':').map((e) => shiftEnd(e, rows, cols)).join(':') + spill
}

/** A formula with its relative references moved by `rows` and `cols`, as a copied formula's are; text in the scan's locale. */
export function shiftFormula(formula: string, rows: number, cols: number, scan: ScanOptions): string {
  if (!rows && !cols) return formula
  const refs = scanReferences(formula, scan).filter(isGridRef)
  let out = formula
  for (const r of [...refs].reverse()) out = out.slice(0, r.at) + shiftRef(r.text, rows, cols) + out.slice(r.end)
  return out
}

/** Whether a formula has a relative reference, which Excel does not allow in the thresholds of color scales, data bars and icon sets. */
export function hasRelative(formula: string, scan: ScanOptions): boolean {
  return scanReferences(formula, scan)
    .filter(isGridRef)
    .some((r) =>
      r.text
        .slice(r.text.lastIndexOf('!') + 1)
        .replace(/#$/, '')
        .split(':')
        .some((e) => {
          const m = END.exec(e)
          return !!m && ((!!m[2] && !m[1]) || (!!m[4] && !m[3]))
        }),
    )
}

/** A rule's own formulas (not its thresholds, which are absolute) passed through `fn`. */
export function mapRuleFormulas(rule: ConditionalFormat, fn: (f: string) => string): ConditionalFormat {
  switch (rule.type) {
    case 'cellIs':
      return { ...rule, formulas: rule.formulas.map(fn) }
    case 'expression':
      return { ...rule, formula: fn(rule.formula) }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
    case 'timePeriod':
      return rule.formula === undefined ? rule : { ...rule, formula: fn(rule.formula) }
    default:
      return rule
  }
}

/** A rule with its formulas rewritten for a first cell `rows` and `cols` away, so each cell keeps its condition. */
export const shiftRule = (rule: ConditionalFormat, rows: number, cols: number, scan: ScanOptions): ConditionalFormat =>
  rows || cols ? mapRuleFormulas(rule, (f) => shiftFormula(f, rows, cols, scan)) : rule

/** What a value box gives: the formula to store (in the locale, without `=`), or the reason it cannot be stored. */
export type BoxFormula = { ok: true; formula: string } | { ok: false; message?: string }

/**
 * A value box's text as Excel's dialogs store it: a formula after `=` (checked as a cell's), a number, date, time or
 * percentage as its number, TRUE or FALSE, otherwise text in quotes. Empty text is refused.
 */
export function boxFormula(text: string, entry: Entry, locale: Locale | undefined): BoxFormula {
  const t = text.trim()
  if (!t) return { ok: false }
  const got = entry.parseEntry(t, locale ? { locale } : {})
  if (!got.ok) return { ok: false, message: got.message }
  if (t.startsWith('=')) return { ok: true, formula: t.slice(1) }
  const v = got.input?.value
  if (typeof v === 'number') return { ok: true, formula: String(v).replace('.', locale?.decimalSeparator ?? '.') }
  if (typeof v === 'boolean') return { ok: true, formula: v ? (locale?.booleans.TRUE ?? 'TRUE') : (locale?.booleans.FALSE ?? 'FALSE') }
  return { ok: true, formula: `"${String(v ?? t).replace(/"/g, '""')}"` }
}

/** A function's name in the locale. */
const fn = (locale: Locale | undefined, name: string) => locale?.functions[name] ?? name

/**
 * The formula Excel writes for a text rule whose text is a reference or a formula (`=$A$1`), over the cell `at`: SEARCH for
 * contains and does not contain, LEFT and RIGHT for begins and ends with. A text constant needs none: xlcalc writes Excel's.
 */
export function textRuleFormula(type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith', ref: string, at: string, locale: Locale | undefined): string {
  const sep = locale?.listSeparator ?? ','
  const search = `${fn(locale, 'ISERROR')}(${fn(locale, 'SEARCH')}(${ref}${sep}${at}))`
  if (type === 'containsText') return `${fn(locale, 'NOT')}(${search})`
  if (type === 'notContainsText') return search
  return `${fn(locale, type === 'beginsWith' ? 'LEFT' : 'RIGHT')}(${at}${sep}${fn(locale, 'LEN')}(${ref}))=${ref}`
}

/** The scan options of a locale: stored text without one. */
export const scanOf = (locale: Locale | undefined): ScanOptions => (locale ? { locale } : {})
