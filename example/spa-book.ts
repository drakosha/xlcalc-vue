// The example app's book helpers: a new empty book, the toolbar's number formats in a locale (the library knows the
// codes), a source whose changes are timed for the status bar, and a file to download.
import { builtinFormatCode, LOCALES, parseEntry, type BookInput, type NumberFormat } from 'xlcalc'
import type { NumberKind, ToolbarNumberFormat } from '../src'
import type { LocalSource } from '../src/local'
import type { BookLocale } from './spa-texts'

/** Excel's first sheet of a new book, named in the book's language. */
export function emptyBook(locale: BookLocale): BookInput {
  return { locale, sheets: [{ name: locale === 'ru-RU' ? 'Лист1' : 'Sheet1', cells: {} }] }
}

/** Excel's Time: the locale's long time (`[$-F400]`), the code Excel writes for it in every locale. */
const LONG_TIME = '[$-F400]h:mm:ss AM/PM'

/** The locale's money with two decimals as Excel sets it for an amount typed with the currency, read by `parseEntry`. */
function moneyFormat(locale: string): string | undefined {
  const regional = LOCALES.get(locale)?.regional
  const c = regional?.currency ?? { symbol: '$', before: true, space: '' }
  const amount = `1${regional?.decimal ?? '.'}50`
  const got = parseEntry(c.before ? `${c.symbol}${c.space}${amount}` : `${amount}${c.space}${c.symbol}`, { locale })
  const format = got.ok ? got.input?.numberFormat : undefined
  return typeof format === 'string' ? format : undefined
}

/**
 * The number formats of Excel's list the toolbar offers, in a loaded locale: General, Number, Percentage, Short Date and
 * Text by the built-in format's number, Currency as typed money gets it, Time as the long time.
 */
export function numberFormatsIn(locale: string): ToolbarNumberFormat[] {
  const offer = (kind: NumberKind, format: NumberFormat): ToolbarNumberFormat[] => {
    const code = typeof format === 'number' ? builtinFormatCode(format, locale) : format
    return code === undefined ? [] : [{ kind, format, code }]
  }
  const money = moneyFormat(locale)
  return [
    ...offer('general', 0),
    ...offer('number', 2),
    ...(money ? offer('currency', money) : []),
    ...offer('percent', 9),
    ...offer('date', 14),
    ...offer('time', LONG_TIME),
    ...offer('text', 49),
  ]
}

/** Methods that change the book: the status bar shows a recomputation while one is out. */
const CHANGES = ['edit', 'paste', 'fill', 'undo', 'redo', 'structure', 'layout', 'format', 'freeze'] as const

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
