// The example's conditional formatting sheet: the rules over synthetic numbers, calculated by xlcalc, written as the
// in-memory workbook's cells (example/cf-book.ts) so the demo shows what the library calculates; `--xlsx <file>` also writes
// the workbook for Excel. Run: node scripts/cf-demo.ts [--xlsx <file>] (Node 22.18+ strips the types).
import { writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createSession, registerLocale, type BookInput, type CellInput, type CellStyle, type ConditionalFormat, type IconSetName } from 'xlcalc'
import { cellAddress } from 'xlcalc/a1'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { writeXlsx } from 'xlcalc/xlsx'
import pkg from 'xlcalc/package.json' with { type: 'json' }

const SHEET = 'Условное форматирование'
const HEAD: CellStyle = { font: { bold: true }, fill: { color: '#E2EFDA' }, border: { bottom: { style: 'thin' } } }
/** The cells' own style under column A's rules: a rule's field replaces it, the rest shows through. */
const OWN: CellStyle = { font: { bold: true, color: '#7030A0' }, border: { bottom: { style: 'thin', color: '#BFBFBF' } } }
const SETS: [IconSetName, number][] = [
  ['3Arrows', 3], ['3ArrowsGray', 3], ['3Flags', 3], ['3TrafficLights1', 3], ['3TrafficLights2', 3], ['3Signs', 3], ['3Symbols', 3],
  ['3Symbols2', 3], ['3Stars', 3], ['3Triangles', 3], ['4Arrows', 4], ['4ArrowsGray', 4], ['4RedToBlack', 4], ['4Rating', 4],
  ['4TrafficLights', 4], ['5Arrows', 5], ['5ArrowsGray', 5], ['5Rating', 5], ['5Quarters', 5], ['5Boxes', 5],
]

const styles: CellStyle[] = [{}, HEAD, OWN]
const cells: Record<string, CellInput> = {}
const rules: ConditionalFormat[] = []
const column = (col: string, title: string, values: number[], style?: number) => {
  cells[`${col}1`] = { value: title, style: 1 }
  values.forEach((v, i) => (cells[`${col}${i + 2}`] = { value: v, ...(style ? { style } : {}) }))
  return `${col}2:${col}${values.length + 1}`
}
const percent = (n: number) => Array.from({ length: n }, (_, i) => ({ value: { type: 'percent' as const, value: String(Math.round((i * 100) / n)) } }))

// A: rules with styles over the cells' own: a fill and a font color, a fill with bold removed and italic added, a border with underline.
const a = column('A', 'Стиль правила', [12, 87, 45, 63, 5, 99, 30, 71, 50, 18], 2)
rules.push(
  { type: 'cellIs', ranges: [a], operator: 'greaterThan', formulas: ['60'], style: { fill: { color: '#FFC7CE' }, font: { color: '#9C0006' } } },
  { type: 'cellIs', ranges: [a], operator: 'lessThan', formulas: ['20'], style: { fill: { color: '#C6EFCE' }, font: { color: '#006100', bold: false, italic: true } } },
  { type: 'expression', ranges: [a], formula: 'OR(A2=50,A2=30)', style: { font: { underline: 'single', strike: true }, border: { top: { style: 'thin', color: '#2F5597' }, right: { style: 'thin', color: '#2F5597' }, bottom: { style: 'thin', color: '#2F5597' }, left: { style: 'thin', color: '#2F5597' } } } },
)
const scale = [3, 18, 27, 41, 50, 59, 66, 78, 85, 97]
rules.push({ type: 'colorScale', ranges: [column('B', 'Шкала', scale)], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'percentile', value: '50' }, color: '#FFEB84' }, { value: { type: 'max' }, color: '#63BE7B' }] })
const positive = [10, 25, 40, 55, 70, 85, 100, 15, 60, 35]
rules.push(
  { type: 'dataBar', ranges: [column('C', 'Градиент', positive)], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', border: '#638EC6' },
  { type: 'dataBar', ranges: [column('D', 'С осью', [-40, -25, -10, 0, 10, 25, 40, 55, 70, 85])], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#63C384', border: '#63C384', negativeColor: '#FF0000', negativeBorder: '#FF0000' },
  { type: 'dataBar', ranges: [column('E', 'Сплошная', positive)], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#FFB628', gradient: false },
  { type: 'dataBar', ranges: [column('F', 'Справа налево', positive)], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#D6007B', border: '#D6007B', direction: 'rightToLeft' },
  { type: 'dataBar', ranges: [column('G', 'Без чисел', positive)], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#008AEF', border: '#008AEF', showValue: false },
  { type: 'iconSet', ranges: [column('H', 'Стрелки', scale)], icons: '3Arrows', points: percent(3) },
  { type: 'iconSet', ranges: [column('I', 'Светофор', scale)], icons: '3TrafficLights1', points: percent(3) },
  { type: 'iconSet', ranges: [column('J', 'Только значок', scale)], icons: '5Rating', points: percent(5), showValue: false },
)

// Every icon set from row 13: its name, then 1 to n, each number on the threshold of its icon.
cells.A13 = { value: 'Набор значков', style: 1 }
for (let i = 1; i <= 5; i++) cells[cellAddress(13, i + 1)] = { value: i, style: 1 }
SETS.forEach(([set, n], k) => {
  const row = 14 + k
  cells[`A${row}`] = { value: set }
  for (let i = 1; i <= n; i++) cells[cellAddress(row, i + 1)] = { value: i }
  const points = Array.from({ length: n }, (_, i) => ({ value: { type: 'num' as const, value: String(i + 1) } }))
  rules.push({ type: 'iconSet', ranges: [`B${row}:${cellAddress(row, n + 1)}`], icons: set, points })
})

const widths: Record<string, number> = { A: 16, B: 9, C: 12, D: 12, E: 12, F: 13, G: 12, H: 10, I: 10, J: 13 }
const book: BookInput = {
  locale: 'ru-RU',
  styles,
  sheets: [{ name: SHEET, cells, conditionalFormats: rules, columns: Object.fromEntries(Object.entries(widths).map(([c, width]) => [c, { width }])) }],
}

registerLocale(ruRU)
const at = process.argv.indexOf('--xlsx')
if (at > 0) writeFileSync(process.argv[at + 1]!, (await writeXlsx({ book })).data)

const v = createSession(book).getViewport(SHEET, 'A1:J33')
const out: Record<string, unknown> = {}
for (const [addr, c] of Object.entries(v.cells)) {
  // style 0 is the workbook's empty one
  const style = c.effectiveStyle ? v.styles[c.effectiveStyle] : undefined
  out[addr] = { value: c.value, text: c.text, ...(c.color ? { color: c.color } : {}), ...(style ? { style } : {}), ...(c.appearance ? { appearance: c.appearance } : {}) }
}
/** A value as the repository's code style writes it: single quotes, unquoted keys where possible. */
const lit = (x: unknown): string => {
  if (typeof x === 'string') return `'${x.replace(/\\/g, '\\\\').replace(/'/g, "\\'")}'`
  if (Array.isArray(x)) return `[${x.map(lit).join(', ')}]`
  if (x && typeof x === 'object') return `{ ${Object.entries(x).map(([k, y]) => `${/^[A-Za-z_$][\w$]*$/.test(k) ? k : lit(k)}: ${lit(y)}`).join(', ')} }`
  return String(x)
}
const file = `// Generated by scripts/cf-demo.ts: the sheet "${SHEET}" with what xlcalc ${pkg.version} calculates for its rules. Do not edit by hand.
import type { CellStyle, CellValue, ConditionalAppearance } from 'xlcalc/contract'

export interface CfCell {
  value: CellValue
  text: string
  color?: string
  style?: CellStyle
  appearance?: ConditionalAppearance
}

export const CF_SHEET = ${lit(SHEET)}
export const CF_WIDTHS: Record<string, number> = ${lit(widths)}
export const CF_CELLS: Record<string, CfCell> = {
${Object.entries(out).map(([addr, c]) => `  ${addr}: ${lit(c)},`).join('\n')}
}
`
writeFileSync(resolve(import.meta.dirname, '../example/cf-book.ts'), file)
console.log(`example/cf-book.ts: ${Object.keys(out).length} cells`)

