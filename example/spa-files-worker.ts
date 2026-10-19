// The app's file thread (see spa-files.ts): reads an xlsx file and sends the workbook back in slices, or receives a snapshot
// in slices and writes it into an xlsx file, whose bytes it transfers. One job per thread.
import type { BookInput } from 'xlcalc'
import { writeXlsx, type XlsxCached } from 'xlcalc/xlsx'
import { loadLocale } from '../src/local/locale'
import { openXlsx } from '../src/local/xlsx'
import { slicesOf, take, withoutCells, type FromThread, type ToThread } from './spa-files'

const post = (m: FromThread, transfer: Transferable[] = []) => postMessage(m, { transfer })
/** The snapshot to write, as its slices arrive; its values in the shape of a file's stored values, which is all writeXlsx reads of a result. */
let snapshot: { book: BookInput; values: XlsxCached['sheets'] } | null = null

async function run(m: ToThread): Promise<void> {
  if (m.kind === 'read') {
    const { book, cached, issues } = await openXlsx(m.file, { locale: m.locale })
    post({ kind: 'head', book: { ...book, sheets: withoutCells(book.sheets) }, cached: { sheets: withoutCells(cached.sheets) }, issues })
    for (const s of slicesOf(book.sheets)) post({ kind: 'cells', part: 'book', ...s })
    for (const s of slicesOf(cached.sheets)) post({ kind: 'cells', part: 'cached', ...s })
    post({ kind: 'read' })
  } else if (m.kind === 'snapshot') snapshot = { book: m.book, values: m.values as XlsxCached['sheets'] }
  else if (m.kind === 'cells') take(m.part === 'book' ? snapshot!.book.sheets : snapshot!.values, m)
  else {
    // the writer parses the workbook in its locale, which this thread loads as the page did for the session
    if (snapshot!.book.locale) await loadLocale(snapshot!.book.locale)
    const { data } = await writeXlsx({ book: snapshot!.book, result: { sheets: snapshot!.values } })
    post({ kind: 'written', data }, [data.buffer])
  }
}

onmessage = (e: MessageEvent<ToThread>) => {
  run(e.data).catch((err: unknown) => post({ kind: 'error', message: err instanceof Error ? err.message : String(err) }))
}
