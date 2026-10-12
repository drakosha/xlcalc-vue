// What the app takes from a file it opens: the book's language, the sheet to show first and what the file lost, by kind.
import type { BookInput, CellInput } from 'xlcalc'
import { quoteSheetIfNeeded } from 'xlcalc/text'
import type { XlsxIssue, XlsxIssueCode } from 'xlcalc/xlsx'
import { BOOK_LOCALES, type BookLocale, type SpaLang } from './spa-texts'

/** The locale of the interface's language: a new book's, and a file's that names none. */
export const localeOfLang = (lang: SpaLang): BookLocale => (lang === 'ru' ? 'ru-RU' : 'en-US')

/**
 * The locale a file's book opens in: the one the file names (xlcalc keeps it in the file), else the interface's, as
 * Excel shows a file's formulas in its user's language; xlsx itself keeps no language.
 */
export function bookLocaleOf(file: string | undefined, lang: SpaLang): BookLocale {
  return BOOK_LOCALES.includes(file as BookLocale) ? (file as BookLocale) : localeOfLang(lang)
}

const hasData = (c: CellInput) => c.formula !== undefined || (c.value !== undefined && c.value !== null && c.value !== '')

/**
 * The sheet to show first: the first visible one with a value or a formula. Excel opens the sheet active when the file
 * was saved, which the library does not read; without it a divider sheet at the front would open empty.
 */
export function firstShownSheet(book: BookInput): string | undefined {
  return book.sheets.find((s) => !s.hidden && Object.values(s.cells).some(hasData))?.name
}

/** Kinds by what they mean to a reader: what changes values first, then content, then the look, settings and names last. */
const ORDER: XlsxIssueCode[] = [
  'externalLink', 'userFunction', 'formula', 'dataTable', 'errorValue', 'errorConstant', 'part', 'table', 'pivotTable',
  'drawing', 'sheetKind', 'cellMetadata', 'connection', 'macros', 'conditionalFormat', 'dataValidation', 'hyperlink',
  'comment', 'richText', 'style', 'numberFormat', 'defaultFont', 'defaultWidth', 'frozenPanes', 'autoFilter', 'outline',
  'sparkline', 'print', 'protection', 'calculation', 'hidden', 'locale', 'name', 'meta', 'textLength', 'noCachedValues',
]
const EXAMPLES = 5

/** Places of one kind: how many, and the first few. */
export interface LossPlaces {
  count: number
  examples: string[]
}
/** What a file lost of one kind. */
export interface LossKind {
  code: XlsxIssueCode
  /** The issues of the kind. */
  count: number
  /** Their messages, each once, with how many issues have it, in the order met. */
  messages: { text: string; count: number }[]
  /** The sheets they are about, in the order met. */
  sheets: string[]
  /** Their cells (`Sheet!A1`, the sheet quoted where Excel quotes it) and names. */
  cells: LossPlaces
  names: LossPlaces
}

/** A place counted once, however many issues of the kind are about it. */
function place(p: LossPlaces, seen: Set<string>, text: string): void {
  if (seen.has(text)) return
  seen.add(text)
  if (p.examples.length < EXAMPLES) p.examples.push(text)
  p.count++
}

/** The reader's issues by kind, in the order of `ORDER`: for the notice of what a file lost and its details. */
export function lossesOf(issues: XlsxIssue[]): LossKind[] {
  type Building = LossKind & { byText: Map<string, { text: string; count: number }>; cellsSeen: Set<string>; namesSeen: Set<string> }
  const kinds = new Map<XlsxIssueCode, Building>()
  for (const i of issues) {
    let k = kinds.get(i.code)
    if (!k) {
      const empty = (): LossPlaces => ({ count: 0, examples: [] })
      kinds.set(i.code, (k = { code: i.code, count: 0, messages: [], byText: new Map(), sheets: [], cells: empty(), names: empty(), cellsSeen: new Set(), namesSeen: new Set() }))
    }
    k.count++
    let m = k.byText.get(i.message)
    if (!m) {
      k.byText.set(i.message, (m = { text: i.message, count: 0 }))
      k.messages.push(m)
    }
    m.count++
    if (i.sheet !== undefined && !k.sheets.includes(i.sheet)) k.sheets.push(i.sheet)
    for (const c of i.cells ?? []) place(k.cells, k.cellsSeen, i.sheet !== undefined ? `${quoteSheetIfNeeded(i.sheet)}!${c}` : c)
    if (i.name !== undefined) place(k.names, k.namesSeen, i.name)
  }
  const rank = (code: XlsxIssueCode) => (ORDER.includes(code) ? ORDER.indexOf(code) : ORDER.length)
  return [...kinds.values()].sort((a, b) => rank(a.code) - rank(b.code)).map(({ code, count, messages, sheets, cells, names }) => ({ code, count, messages, sheets, cells, names }))
}
