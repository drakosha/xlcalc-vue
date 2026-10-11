// An xlsx file into a book and a book with its values into a file, through `xlcalc/xlsx`, which loads only when used.
import type { XlsxIssue, XlsxRead, XlsxWritten } from 'xlcalc/xlsx'
import type { LocalSource } from './types'

/** Read a file (the bytes, or a `File` from an input or a drop) into a book; `issues` lists what the book cannot hold. */
export async function openXlsx(data: Blob | ArrayBuffer | Uint8Array): Promise<XlsxRead> {
  const bytes = data instanceof Blob ? new Uint8Array(await data.arrayBuffer()) : data
  const { readXlsx } = await import('xlcalc/xlsx')
  return readXlsx(bytes)
}

/** Write the source's book with its computed values (`formulas`, the default, or `values` only); `issues` lists what the file cannot hold. */
export async function saveXlsx(source: LocalSource, opts: { mode?: 'formulas' | 'values' } = {}): Promise<XlsxWritten> {
  const [{ writeXlsx }, snapshot] = await Promise.all([import('xlcalc/xlsx'), source.snapshot()])
  return writeXlsx(snapshot, opts.mode ? { mode: opts.mode } : {})
}

export interface IssueGroup {
  code: XlsxIssue['code']
  count: number
  /** The sheets the issues are about, in the order met. */
  sheets: string[]
  /** The first issue's message. */
  message: string
}

/** Issues by kind, in the order their kinds first appear: for a notice of what a file lost. */
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
