// An xlsx file into a book and a book with its values into a file, through `xlcalc/xlsx`, which loads only when used.
import type { Locale } from 'xlcalc'
import type { XlsxIssue, XlsxRead, XlsxWritten } from 'xlcalc/xlsx'
import { loadLocale } from './locale'
import type { LocalSource } from './types'

/** The language of the issue messages: a tag (`ru-RU`, whose module is loaded on demand) or a locale object; en-US by default. */
export interface XlsxIssuesOptions {
  locale?: string | Locale
}

/** The reader's and writer's `locale`, a tag's module loaded first: the library takes a tag it has loaded and refuses others. */
async function issuesIn(opts: XlsxIssuesOptions): Promise<{ locale?: string | Locale }> {
  if (opts.locale === undefined) return {}
  if (typeof opts.locale === 'string') await loadLocale(opts.locale)
  return { locale: opts.locale }
}

/** Reads a file (the bytes, or a `File` from a file input or a drop) into a workbook; `issues` lists what the workbook cannot keep, in `locale`'s language. */
export async function openXlsx(data: Blob | ArrayBuffer | Uint8Array, opts: XlsxIssuesOptions = {}): Promise<XlsxRead> {
  const bytes = data instanceof Blob ? new Uint8Array(await data.arrayBuffer()) : data
  const [{ readXlsx }, read] = await Promise.all([import('xlcalc/xlsx'), issuesIn(opts)])
  return readXlsx(bytes, read)
}

/** Writes the source's workbook with its calculated values (`formulas`, the default, or `values` only); `issues` lists what the file cannot keep, in `locale`'s language. */
export async function saveXlsx(source: LocalSource, opts: XlsxIssuesOptions & { mode?: 'formulas' | 'values' } = {}): Promise<XlsxWritten> {
  const [{ writeXlsx }, snapshot, write] = await Promise.all([import('xlcalc/xlsx'), source.snapshot(), issuesIn(opts)])
  return writeXlsx(snapshot, { ...(opts.mode ? { mode: opts.mode } : {}), ...write })
}

export interface IssueGroup {
  code: XlsxIssue['code']
  count: number
  /** The sheets the issues concern, in order of appearance. */
  sheets: string[]
  /** The first issue's message. */
  message: string
}

/** Issues grouped by kind, in order of first appearance: for a notice of what a file lost. */
export function groupIssues(issues: XlsxIssue[]): IssueGroup[] {
  const groups = new Map<string, IssueGroup>()
  for (const i of issues) {
    let g = groups.get(i.code)
    if (!g) groups.set(i.code, (g = { code: i.code, count: 0, sheets: [], message: i.message }))
    g.count++
    if (i.sheet && !g.sheets.includes(i.sheet)) g.sheets.push(i.sheet)
  }
  return [...groups.values()]
}
