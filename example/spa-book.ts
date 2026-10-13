// The example app's book helpers: a new empty book, a source whose changes are timed for the status bar, and a file to
// download.
import type { BookInput } from 'xlcalc'
import type { LocalSource } from '../src/local'
import type { BookLocale } from './spa-texts'

/** Excel's first sheet of a new book, named in the book's language. */
export function emptyBook(locale: BookLocale): BookInput {
  return { locale, sheets: [{ name: locale === 'ru-RU' ? 'Лист1' : 'Sheet1', cells: {} }] }
}

/** Methods that change the book: the status bar shows a recomputation while one is out. */
const CHANGES = ['edit', 'paste', 'fill', 'pick', 'undo', 'redo', 'structure', 'layout', 'format', 'replace', 'outline', 'freeze'] as const

/** The source with its changes timed: `begin` as one goes out, `end` with its milliseconds as it comes back. */
export function timed(source: LocalSource, begin: () => void, end: (ms: number) => void): LocalSource {
  // the source's other members (ready, locale, history's getter) come through the prototype
  const out = Object.create(source) as LocalSource
  for (const name of CHANGES) {
    const call = source[name] as (...args: unknown[]) => Promise<unknown>
    const wrapped = async (...args: unknown[]) => {
      begin()
      const t0 = performance.now()
      try {
        return await call(...args)
      } finally {
        end(performance.now() - t0)
      }
    }
    Object.defineProperty(out, name, { value: wrapped, enumerable: true })
  }
  return out
}

/** Hand bytes to the browser as a download named `name`. */
export function download(data: Uint8Array, name: string): void {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(new Blob([data as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(a.href), 1000)
}
