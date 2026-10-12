// A synthetic ru-RU book for the local source: a showcase sheet (a table, a note, links, formats, conditional formats) and xlcalc's synthetic book.
import type { BookInput, CellInput, CellStyle, ConditionalFormat } from 'xlcalc'
import { synthBook } from 'xlcalc/unstable/synth'

const HEAD: CellStyle = { font: { bold: true, color: '#ffffff' }, fill: { color: '#2f5597' }, alignment: { horizontal: 'center' } }
const TITLE: CellStyle = { font: { bold: true, size: 14 }, alignment: { horizontal: 'center', vertical: 'center' } }
const TOTAL: CellStyle = { font: { bold: true }, border: { top: { style: 'thin' }, bottom: { style: 'double' } } }
/** A data bar over the amounts and a colour scale over the shares: an amount typed recolours every share. */
const RULES: ConditionalFormat[] = [
  { type: 'dataBar', ranges: ['B3:B7'], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', border: '#638EC6' },
  { type: 'colorScale', ranges: ['D3:D7'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'percentile', value: '50' }, color: '#FFEB84' }, { value: { type: 'max' }, color: '#63BE7B' }] },
]

/** The synthetic book of `rows` data rows, with a showcase sheet first and the data sheet's header frozen and formatted. */
export function localBook(rows: number): BookInput {
  const { book } = synthBook({ seed: 7, rows })
  const styles: CellStyle[] = [{}, HEAD, TITLE, TOTAL]
  const items = ['Чай', 'Кофе', 'Сахар', 'Молоко', 'Печенье']
  const cells: Record<string, CellInput> = {
    A1: { value: 'Витрина: таблица, примечание, ссылки, форматы', style: 2 },
    A2: { value: 'Товар', style: 1 },
    B2: { value: 'Сумма', style: 1 },
    C2: { value: 'Двойная', style: 1 },
    D2: { value: 'Доля', style: 1 },
    E2: { value: 'Дата', style: 1 },
  }
  items.forEach((name, i) => {
    const r = i + 3
    cells[`A${r}`] = { value: name }
    cells[`B${r}`] = { value: (i + 1) * 1250.5, numberFormat: '#,##0.00' }
    cells[`C${r}`] = { formula: 'Продажи[[#This Row],[Сумма]]*2', numberFormat: '#,##0.00' }
    cells[`D${r}`] = { formula: `B${r}/$B$9`, numberFormat: '0.0%' }
    cells[`E${r}`] = { value: { date: `2026-0${i + 1}-15` }, numberFormat: 14 }
  })
  cells.A9 = { value: 'Итого', style: 3 }
  cells.B9 = { formula: 'SUM(Продажи[Сумма])', numberFormat: '#,##0.00', style: 3 }
  cells.C9 = { formula: 'SUM(Продажи[Двойная])', numberFormat: '#,##0.00', style: 3 }
  cells.A3 = { value: 'Чай', note: { author: 'Аня', text: 'Синтетика: цифры выдуманы' } }
  cells.A11 = { value: 'Сайт библиотеки', link: { url: 'https://github.com/drakosha/xlcalc', tooltip: 'xlcalc на GitHub' } }
  cells.A12 = { value: 'К данным', link: { location: "'Данные'!A1" } }
  const showcase = {
    name: 'Витрина',
    cells,
    columns: { A: { width: 22 }, 'B:C': { width: 14 }, E: { width: 12 } },
    rows: { '1': { height: 30 } },
    merges: ['A1:E1'],
    frozen: { rows: 2 },
    conditionalFormats: RULES,
    tables: [
      {
        name: 'Продажи',
        range: 'A2:E7',
        columns: [{ name: 'Товар' }, { name: 'Сумма' }, { name: 'Двойная', formula: 'Продажи[[#This Row],[Сумма]]*2' }, { name: 'Доля' }, { name: 'Дата' }],
      },
    ],
  }
  const sheets = book.sheets.map((s) =>
    s.name === 'Данные'
      ? { ...s, frozen: { rows: 1 }, columns: { A: { numberFormat: 14, width: 11 }, B: { width: 14 }, D: { numberFormat: '#,##0.00', width: 13 }, J: { numberFormat: 14, width: 11 } }, rows: { '1': { style: 1 } } }
      : s,
  )
  return { ...book, locale: 'ru-RU', styles, sheets: [showcase, ...sheets] }
}
