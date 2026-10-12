// The values the app saved in a file against the values Excel saved in its copy after a full recalculation: every cell with
// a value in either file (a value in one only differs), cells of volatile formulas left out (they differ by the clock).
// Run: SAVED=app.xlsx EXCEL=excel.xlsx npm run check:compare-values (node 22.18+ strips the types; a large book wants
// NODE_OPTIONS=--max-old-space-size=12288). Exits 1 when a value differs.
import { readFile } from 'node:fs/promises'
import { createSession, registerLocale, type CellValue } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { readXlsx } from 'xlcalc/xlsx'

registerLocale(ruRU)
const saved = process.env.SAVED
const excel = process.env.EXCEL
if (!saved || !excel) throw new Error('SAVED: the file the app saved; EXCEL: its copy Excel saved after CalculateFull')
const SHOWN = Number(process.env.SHOWN ?? 10)

const isError = (v: CellValue | undefined): v is Extract<CellValue, { error: unknown }> => !!v && typeof v === 'object' && 'error' in v
/** Numbers within 1e-12 or a relative 1e-14 (dates are serial numbers), text exactly, booleans and errors by kind and code. */
function same(a: CellValue | undefined, b: CellValue | undefined): boolean {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-12 || Math.abs(a - b) <= 1e-14 * Math.max(Math.abs(a), Math.abs(b))
  if (isError(a) || isError(b)) return isError(a) && isError(b) && a.error === b.error
  return a === b
}
const empty = (v: CellValue | undefined) => v === undefined || v === null || v === ''

const ours = await readXlsx(await readFile(saved), { locale: 'ru-RU' })
const theirs = await readXlsx(await readFile(excel), { locale: 'ru-RU' })
// Volatile cells by the session over the file the app saved.
const volatile = new Set(createSession(ours.book, { limits: { cells: 50_000_000 } }).getResult().volatile.map((c) => `${c.sheet}!${c.cell}`))

/** A sheet's values by address: the cells' own values and what formulas and spills computed, as each file stores them. */
function valuesOf(file: typeof ours, i: number): Map<string, CellValue> {
  const out = new Map<string, CellValue>()
  for (const [addr, cell] of Object.entries(file.book.sheets[i]?.cells ?? {})) {
    const v = cell.value
    // A date typed into a book is { date }; a file read back gives serial numbers, so such a value shows as its text.
    if (cell.formula === undefined && v !== undefined && !empty(v as CellValue)) out.set(addr, typeof v === 'object' && v !== null && 'date' in v ? v.date : v)
  }
  for (const [addr, cell] of Object.entries(file.cached?.sheets?.[i]?.cells ?? {})) if (!empty(cell.value)) out.set(addr, cell.value)
  return out
}

let compared = 0
let skipped = 0
const differ: string[] = []
const names = ours.book.sheets.map((s) => s.name)
const missing = theirs.book.sheets.map((s) => s.name).filter((n) => !names.includes(n))
if (missing.length) differ.push(`sheets only in Excel's file: ${missing.join(', ')}`)
names.forEach((name, i) => {
  const j = theirs.book.sheets.findIndex((s) => s.name === name)
  if (j < 0) return void differ.push(`sheet ${name} is not in Excel's file`)
  const a = valuesOf(ours, i)
  const b = valuesOf(theirs, j)
  for (const addr of new Set([...a.keys(), ...b.keys()])) {
    if (volatile.has(`${name}!${addr}`)) {
      skipped++
      continue
    }
    compared++
    if (!same(a.get(addr), b.get(addr))) differ.push(`${name}!${addr}: app ${JSON.stringify(a.get(addr) ?? null)}, Excel ${JSON.stringify(b.get(addr) ?? null)}`)
  }
})
console.log(`${compared} cells compared (${skipped} volatile left out), ${differ.length} differ`)
for (const line of differ.slice(0, SHOWN)) console.log(`  ${line}`)
if (differ.length) process.exit(1)
