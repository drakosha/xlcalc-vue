// An xlsx file into a workbook and a workbook with its values into a file, through `xlcalc/xlsx`, which loads only when used.
import type { Locale } from 'xlcalc'
import type { XlsxIssue, XlsxWritten } from 'xlcalc/xlsx'
import { loadLocale } from './locale'
import type { LocalSource, OpenedXlsx } from './types'

/** The language of the issue messages: a tag (`ru-RU`, whose module is loaded on demand) or a locale object; en-US by default. */
export interface XlsxIssuesOptions {
  locale?: string | Locale
}

/** The reader's and writer's `locale`, with a tag's module loaded first: the library accepts a tag it has loaded and rejects others. */
async function issuesIn(opts: XlsxIssuesOptions): Promise<{ locale?: string | Locale }> {
  if (opts.locale === undefined) return {}
  if (typeof opts.locale === 'string') await loadLocale(opts.locale)
  return { locale: opts.locale }
}

/** A file's bytes: a `File` from a file input or a drop, or the bytes themselves. */
type FileData = Blob | ArrayBuffer | Uint8Array
const bytesOf = async (data: FileData) => (data instanceof Blob ? new Uint8Array(await data.arrayBuffer()) : data)

/**
 * Reads a file into a workbook; `issues` lists what the workbook cannot keep, in `locale`'s language, and `file` names a
 * `File` for the session (`CELL("filename")`, `INFO`). With `keep`, through xlcalc's `openKept` (0.57.0 or later), so that
 * `saveXlsx` with `over` can write the workbook over this file.
 */
export async function openXlsx(data: FileData, opts: XlsxIssuesOptions & { keep?: boolean } = {}): Promise<OpenedXlsx> {
  const [bytes, read] = await Promise.all([bytesOf(data), issuesIn(opts)])
  const got = opts.keep ? (await import('xlcalc/xlsx/keep')).openKept(bytes, read) : (await import('xlcalc/xlsx')).readXlsx(bytes, read)
  // a browser does not tell a file's folder: the name alone
  const name = typeof File === 'function' && data instanceof File && data.name ? data.name : undefined
  return name ? { ...(await got), file: { name } } : got
}

/**
 * Writes the source's workbook with its calculated values (`formulas`, the default, or `values` only); `issues` lists what
 * the file cannot keep, in `locale`'s language. In Manual the source calculates first, as Excel recalculates a workbook
 * before saving it (its default), so the values the sheet shows are the ones written; the file keeps the option. With `over`, the file the workbook was opened from with `keep`, through
 * xlcalc's `writeKept`: what the workbook does not hold stays as the file has it, and a change the file cannot take is
 * rejected with xlcalc's `KeptRefusal`.
 */
export async function saveXlsx(source: LocalSource, opts: XlsxIssuesOptions & ({ mode?: 'formulas' | 'values'; over?: undefined } | { over: FileData; mode?: undefined }) = {}): Promise<XlsxWritten> {
  await source.calculate()
  const [snapshot, write] = await Promise.all([source.snapshot(), issuesIn(opts)])
  if (opts.over) {
    const [{ writeKept }, original] = await Promise.all([import('xlcalc/xlsx/keep'), bytesOf(opts.over)])
    return writeKept(original, snapshot, write)
  }
  return (await import('xlcalc/xlsx')).writeXlsx(snapshot, { ...(opts.mode ? { mode: opts.mode } : {}), ...write })
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
