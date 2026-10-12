// xlsx files read and written in a thread of their own (spa-files-worker.ts): a large file takes seconds either way,
// and the page keeps drawing meanwhile. Books cross between the threads as their sheets without cells, then the cells
// in slices, so that neither side stops for a whole book at once.
import type { BookInput, BookSnapshot } from 'xlcalc'
import type { XlsxIssue } from 'xlcalc/xlsx'

/** Cells per message: a slice is taken in one task of a few milliseconds. */
export const SLICE = 20_000

type Cells = Record<string, unknown>
/** Which sheets a slice belongs to: the book's, or the result's values. */
export type Part = 'book' | 'result'

/** From the thread: a file read (the book without cells first), a file written, or why not. */
export type FromThread =
  | { kind: 'head'; book: BookInput; issues: XlsxIssue[] }
  | { kind: 'cells'; part: Part; sheet: number; cells: Cells }
  | { kind: 'read' }
  | { kind: 'written'; data: Uint8Array }
  | { kind: 'error'; message: string }
/** To the thread: a file to read, or a snapshot to write (its sheets without cells first, the cells after, then `write`). */
export type ToThread =
  | { kind: 'read'; file: Blob; locale: string }
  | { kind: 'snapshot'; book: BookInput; values: { name: string; cells: Cells }[] }
  | { kind: 'cells'; part: Part; sheet: number; cells: Cells }
  | { kind: 'write' }

/** Sheets with their cells left out: what goes ahead of the slices. */
export const withoutCells = <S extends { cells: Cells }>(sheets: S[]): S[] => sheets.map((s) => ({ ...s, cells: {} }))

/** The cells of sheets in slices of `SLICE`, sheet by sheet. */
export function* slicesOf(sheets: { cells: Cells }[]): Generator<{ sheet: number; cells: Cells }> {
  for (const [sheet, s] of sheets.entries()) {
    let cells: Cells = {}
    let n = 0
    for (const addr in s.cells) {
      cells[addr] = s.cells[addr]
      if (++n === SLICE) {
        yield { sheet, cells }
        cells = {}
        n = 0
      }
    }
    if (n) yield { sheet, cells }
  }
}

/** A slice put back into its sheet. */
export function take(sheets: { cells: Cells }[], m: { sheet: number; cells: Cells }): void {
  Object.assign(sheets[m.sheet]!.cells, m.cells)
}

const failure = (e: ErrorEvent) => new Error(e.message || 'The file thread failed')

/** A turn of the page's event loop without a timer's delay: frames are drawn between the slices it sends. */
function turn(): Promise<void> {
  const { port1, port2 } = new MessageChannel()
  return new Promise((done) => {
    port1.onmessage = () => {
      port1.close()
      done()
    }
    port2.postMessage(0)
  })
}

/** A job of a new thread, which ends with it: `send` gives it its messages, `on` takes its answers until one settles the job. */
function job<T>(send: (post: (m: ToThread, transfer?: Transferable[]) => void) => Promise<void> | void, on: (m: FromThread, done: (v: T) => void, fail: (e: Error) => void) => void): Promise<T> {
  const worker = new Worker(new URL('./spa-files-worker.ts', import.meta.url), { type: 'module' })
  return new Promise<T>((done, fail) => {
    worker.onmessage = (e: MessageEvent<FromThread>) => (e.data.kind === 'error' ? fail(new Error(e.data.message)) : on(e.data, done, fail))
    worker.onerror = (e) => fail(failure(e))
    Promise.resolve(send((m, transfer = []) => worker.postMessage(m, transfer))).catch(fail)
  }).finally(() => worker.terminate())
}

/** Read a file into a book in a thread; `issues` say what the book cannot hold, in `locale`'s language. */
export function readInThread(file: Blob, locale: string): Promise<{ book: BookInput; issues: XlsxIssue[] }> {
  let head: { book: BookInput; issues: XlsxIssue[] } | null = null
  return job(
    (post) => post({ kind: 'read', file, locale }),
    (m, done) => {
      if (m.kind === 'head') head = { book: m.book, issues: m.issues }
      else if (m.kind === 'cells') take(head!.book.sheets as { cells: Cells }[], m)
      else if (m.kind === 'read') done(head!)
    },
  )
}

/** Write a snapshot of a book and its values into an xlsx file in a thread; the slices go a turn of the page apart. */
export function writeInThread(snapshot: BookSnapshot): Promise<Uint8Array> {
  const { book, result } = snapshot
  return job<Uint8Array>(
    async (post) => {
      post({ kind: 'snapshot', book: { ...book, sheets: withoutCells(book.sheets) }, values: withoutCells(result.sheets) })
      for (const [part, sheets] of [['book', book.sheets], ['result', result.sheets]] as const) {
        for (const s of slicesOf(sheets as { cells: Cells }[])) {
          post({ kind: 'cells', part, ...s })
          await turn()
        }
      }
      post({ kind: 'write' })
    },
    (m, done) => {
      if (m.kind === 'written') done(m.data)
    },
  )
}
