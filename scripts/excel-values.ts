// The library's values of a book against the values Excel saved in it: every formula cell compared with readXlsx's
// `cached` after a session opened with them computes the book (a circular reference holds them, as Excel opens the file),
// cells that depend on volatile functions left out (they differ by the clock).
// Run: XLSX=a.xlsx[:b.xlsx] npm run check:excel-values (node 22.18+ strips the types; a large book wants
// NODE_OPTIONS=--max-old-space-size=12288). Exits 1 when a value differs.
import { readFile } from 'node:fs/promises'
import { createSession, registerLocale, sameValue, type CellValue } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { readXlsx } from 'xlcalc/xlsx'

registerLocale(ruRU)
const paths = (process.env.XLSX ?? '').split(':').filter(Boolean)
if (!paths.length) throw new Error('XLSX: the path of a book, several by ":"')
const SHOWN = Number(process.env.SHOWN ?? 5)

const isError = (v: CellValue | undefined): boolean => !!v && typeof v === 'object' && 'error' in v
/** Numbers within a relative 1e-9, everything else as the values hash compares it. */
function same(a: CellValue | undefined, b: CellValue | undefined): boolean {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b))
  return sameValue(a ?? null, b ?? null)
}

let differ = 0
for (const path of paths) {
  const started = performance.now()
  const { book, cached } = await readXlsx(await readFile(path), { locale: 'ru-RU' })
  const read = performance.now()
  const result = createSession(book, { cached, limits: { cells: 50_000_000 } }).getResult()
  const computed = performance.now()
  const volatile = new Set(result.volatile.map((c) => `${c.sheet}!${c.cell}`))
  let compared = 0
  let skipped = 0
  let errors = 0
  let other = 0
  const shown: string[] = []
  book.sheets.forEach((sheet, i) => {
    const ours = result.sheets[i]?.cells ?? {}
    const theirs = cached?.sheets?.[i]?.cells ?? {}
    for (const [addr, cell] of Object.entries(sheet.cells)) {
      if (cell.formula === undefined || !(addr in theirs)) continue
      if (volatile.has(`${sheet.name}!${addr}`)) {
        skipped++
        continue
      }
      compared++
      const a = ours[addr]?.value
      const b = theirs[addr]!.value
      if (same(a, b)) continue
      if (isError(a) && !isError(b)) errors++
      else other++
      if (shown.length < SHOWN) shown.push(`  ${sheet.name}!${addr} =${cell.formula.slice(0, 80)}: ${JSON.stringify(a)}, Excel ${JSON.stringify(b)}`)
    }
  })
  differ += errors + other
  console.log(`${path.split('/').pop()}: read ${((read - started) / 1000).toFixed(1)} s, computed ${((computed - read) / 1000).toFixed(1)} s; ${compared} formulas (${skipped} volatile left out), ${errors + other} differ (${errors} an error where Excel has a value, ${other} other)`)
  for (const line of shown) console.log(line)
}
if (differ) process.exit(1)
