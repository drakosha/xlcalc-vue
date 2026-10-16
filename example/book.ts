// The "all formatting" book: every formatting feature of 0.1.0 on its own sheet, conditional formats, data validation and a chain of formulas; synthetic data only.
import { cellAddress } from 'xlcalc/a1'
import type { BorderLine, CellStyle } from 'xlcalc/contract'
import type { MemoryBook, MemoryCell, MemorySheet } from '../src'
import { CF_CELLS, CF_SHEET, CF_WIDTHS } from './cf-book'
import { listSheet } from './list-book'

const styles: CellStyle[] = [{}]
/** Add a style to the book's table and return its index. */
const st = (style: CellStyle): number => styles.push(style) - 1

const KINDS: BorderLine['style'][] = ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']

function formatting(): MemorySheet {
  const cells: Record<string, MemoryCell> = {}
  const head = st({ font: { bold: true, size: 12 }, fill: { color: '#dbe7f5' }, border: { bottom: { style: 'medium', color: '#2f5597' } } })
  cells.A1 = { value: 'Шрифт', style: head }
  const fonts: [string, CellStyle][] = [
    ['Calibri 11', { font: { name: 'Calibri', size: 11 } }],
    ['Arial 10', { font: { name: 'Arial', size: 10 } }],
    ['Times New Roman 12', { font: { name: 'Times New Roman', size: 12 } }],
    ['Courier New 11', { font: { name: 'Courier New', size: 11 } }],
    ['Кегль 8', { font: { size: 8 } }],
    ['Кегль 16', { font: { size: 16 } }],
    ['Жирный', { font: { bold: true } }],
    ['Курсив', { font: { italic: true } }],
    ['Подчёркнутый', { font: { underline: 'single' } }],
    ['Двойное подчёркивание', { font: { underline: 'double' } }],
    ['Зачёркнутый', { font: { strike: true } }],
    ['Красный текст', { font: { color: '#c00000' } }],
    ['Жёлтая заливка', { fill: { color: '#ffff00' } }],
    ['Белый на синем', { font: { color: '#ffffff', bold: true }, fill: { color: '#2f5597' } }],
  ]
  fonts.forEach(([text, style], i) => (cells[`A${i + 2}`] = { value: text, style: st(style) }))

  cells.C1 = { value: 'Рамки', style: head }
  KINDS.forEach((kind, i) => {
    const line: BorderLine = { style: kind, color: i % 3 === 1 ? '#c00000' : undefined }
    cells[`C${i * 2 + 3}`] = { value: kind, style: st({ border: { top: line, right: line, bottom: line, left: line } }) }
  })

  cells.E1 = { value: 'По горизонтали', style: head }
  const horizontal: [MemoryCell, CellStyle][] = [
    [{ value: 1234.5, text: '1234.5' }, {}],
    [{ value: 'текст' }, {}],
    [{ value: true, text: 'TRUE' }, {}],
    [{ value: { error: 'DIV0' } }, {}],
    [{ value: 'left' }, { alignment: { horizontal: 'left' } }],
    [{ value: 'center' }, { alignment: { horizontal: 'center' } }],
    [{ value: 'right' }, { alignment: { horizontal: 'right' } }],
    [{ value: '*-' }, { alignment: { horizontal: 'fill' } }],
    [{ value: 'justify: длинный текст выравнивается по ширине ячейки' }, { alignment: { horizontal: 'justify' } }],
    [{ value: 'distributed текст' }, { alignment: { horizontal: 'distributed' } }],
    [{ value: 'отступ 1' }, { alignment: { horizontal: 'left', indent: 1 } }],
    [{ value: 'отступ 3' }, { alignment: { horizontal: 'left', indent: 3 } }],
    [{ value: 'справа, отступ 2' }, { alignment: { horizontal: 'right', indent: 2 } }],
  ]
  horizontal.forEach(([cell, style], i) => (cells[`E${i + 2}`] = { ...cell, style: st(style) }))
  cells.E15 = { value: 'по центру выделения', style: st({ alignment: { horizontal: 'centerContinuous' }, fill: { color: '#fde9d9' } }) }
  cells.F15 = { style: st({ alignment: { horizontal: 'centerContinuous' }, fill: { color: '#fde9d9' } }) }

  cells.H1 = { value: 'По вертикали', style: head }
  cells.H2 = { value: 'сверху', style: st({ alignment: { vertical: 'top' }, border: { bottom: { style: 'thin' } } }) }
  cells.H3 = { value: 'по центру', style: st({ alignment: { vertical: 'center' }, border: { bottom: { style: 'thin' } } }) }
  cells.H4 = { value: 'снизу', style: st({ alignment: { vertical: 'bottom' }, border: { bottom: { style: 'thin' } } }) }
  cells.H5 = { value: 'Перенос по словам: длинный текст в высокой строке', style: st({ alignment: { wrap: true, vertical: 'top' } }) }
  cells.H6 = { value: 'Без переноса: длинный текст обрезан краем, соседняя ячейка занята', style: st({}) }
  cells.I6 = { value: 'занято' }
  cells.H8 = { value: 'Строки через Alt+Enter:\nвторая\nтретья\nчетвёртая\nпятая\nшестая', style: st({ alignment: { wrap: true, vertical: 'top' } }) }
  return {
    name: 'Оформление',
    cells,
    widths: { A: 24, B: 3, C: 18, D: 3, E: 22, F: 12, G: 3, H: 22 },
    // Excel stores a taller row for a bigger font; justify needs room for its lines.
    heights: { '1': 16.5, '2': 40, '3': 40, '4': 40, '5': 58, '7': 21, '8': 90, '10': 45 },
  }
}

function formats(): MemorySheet {
  const head = st({ font: { bold: true }, fill: { color: '#e2efda' } })
  const right = st({})
  // Name, value, shown text, the format's code (the text a cell opens with follows it), the section's colour
  const rows: [string, number, string, string, string?][] = [
    ['Число с разрядами', 1234567.891, '1 234 567,89', '#,##0.00'],
    ['Деньги', 1234.5, '1 234,50 ₽', '#,##0.00 "₽"'],
    ['Отрицательное красным', -42, '-42,00', '0.00;[Red]-0.00', '#ff0000'],
    ['Положительное синим', 42, '42,00', '[Blue]0.00', '#0070c0'],
    ['Процент', 0.355, '35,5%', '0.0%'],
    ['Дата', 46096, '15.03.2026', 'dd.mm.yyyy'],
    ['Время', 0.75, '18:00', 'h:mm'],
    ['Экспонента', 0.000012, '1,20E-05', '0.00E+00'],
    ['Дробь', 1.25, '1 1/4', '# ?/?'],
    ['Не влезает в ширину', 123456789012, '########', '0'],
  ]
  const cells: Record<string, MemoryCell> = { A1: { value: 'Формат', style: head }, B1: { value: 'Показано', style: head }, C1: { value: 'Значение', style: head } }
  rows.forEach(([name, value, text, numberFormat, color], i) => {
    cells[`A${i + 2}`] = { value: name }
    cells[`B${i + 2}`] = { value, text, numberFormat, color, style: right }
    cells[`C${i + 2}`] = { value }
  })
  return { name: 'Форматы', cells, widths: { A: 26, B: 16, C: 16 } }
}

function bars(): MemorySheet {
  const cells: Record<string, MemoryCell> = { A1: { value: 'Отдел' }, B1: { value: 'План' }, C1: { value: 'Факт' } }
  const data = [120, 340, 80, 560, 410, 230, 15, 600]
  data.forEach((v, i) => {
    cells[`A${i + 2}`] = { value: `Отдел ${i + 1}` }
    cells[`B${i + 2}`] = { value: v, appearance: { bar: { start: 0, end: v / 600, color: '#638ec6', gradient: true } } }
    cells[`C${i + 2}`] = { value: v * 0.8, appearance: { bar: { start: 0, end: (v * 0.8) / 600, color: '#63be7b', gradient: true } } }
  })
  return { name: 'Гистограммы', cells, widths: { A: 14, B: 20, C: 20 } }
}

/** Conditional formats as xlcalc computed them for the sheet's rules (example/cf-book.ts, made by scripts/cf-demo.ts). */
function conditional(): MemorySheet {
  const ids = new Map<string, number>()
  const cells: Record<string, MemoryCell> = {}
  for (const [addr, { style, ...cell }] of Object.entries(CF_CELLS)) {
    const key = JSON.stringify(style)
    if (style && !ids.has(key)) ids.set(key, st(style))
    cells[addr] = { ...cell, ...(style ? { style: ids.get(key) } : {}) }
  }
  return { name: CF_SHEET, cells, widths: CF_WIDTHS }
}

function links(): MemorySheet {
  return {
    name: 'Ссылки',
    cells: {
      A1: { value: 'Вид ссылки' },
      B1: { value: 'Ссылка' },
      A2: { value: 'Внешний адрес' },
      B2: { value: 'example.org', link: { href: 'https://example.org/', tooltip: 'Пример сайта' } },
      A3: { value: 'Место в книге' },
      B3: { value: 'Длинный лист, строка 500', link: { location: "'Длинный лист'!B500" } },
      A4: { value: 'Лист целиком' },
      B4: { value: 'К закреплению', link: { location: 'Закрепление' } },
      A5: { value: 'Файл у хоста' },
      B5: { value: 'Отчёт.xlsx', link: { target: 'file:42', href: 'Отчёт.xlsx', tooltip: 'Открыть файл в приложении' } },
      A6: { value: 'Небезопасная' },
      B6: { value: 'не откроется', link: { href: 'javascript:alert(1)' } },
      A7: { value: 'Почта' },
      B7: { value: 'mail@example.org', link: { href: 'mailto:mail@example.org' } },
    },
    widths: { A: 18, B: 30 },
  }
}

function notes(): MemorySheet {
  return {
    name: 'Примечания',
    cells: {
      A1: { value: 'Статья' },
      B1: { value: 'Сумма' },
      A2: { value: 'Аренда', note: { author: 'Автор', text: 'По договору от 01.03' } },
      B2: { value: 150000, text: '150 000', note: { text: 'Без автора' } },
      A3: { value: 'Связь' },
      B3: { value: 12000, text: '12 000', note: { author: 'Проверяющий', text: 'Длинное примечание: сумма за квартал, без учёта скидки, пересчитать после закрытия месяца' } },
    },
    widths: { A: 18, B: 14 },
  }
}

function formulas(): MemorySheet {
  const cells: Record<string, MemoryCell> = { A1: { value: 'Значение' }, B1: { value: 'Формула' } }
  for (let r = 2; r <= 6; r++) cells[`A${r}`] = { value: r * 10 }
  cells.B2 = { value: 200, formula: 'СУММ(A2:A6)', note: { author: 'Автор', text: 'Итог столбца' } }
  cells.B3 = { value: 40, formula: 'СРЗНАЧ(A2:A6)-A2+A3' }
  cells.B4 = { value: 30, formula: "ЕСЛИ(A2>0,5;A3;'Ссылки'!A1)" }
  cells.B5 = { value: 1.5, text: '1,5', formula: 'ОКРУГЛ(A2/A3*1,5;2)' }
  cells.B6 = { value: 60, formula: 'СУММ(A:A)-СУММ(A2:A5)+Форматы!C2*0' }
  return { name: 'Формулы', cells, widths: { A: 12, B: 18 } }
}

/** A chain of formulas for the precedents' panel: profit from price and quantity, a rate beside, a plan from another sheet; values as the formulas give them. */
function chain(): MemorySheet {
  const head = st({ font: { bold: true }, fill: { color: '#d9e1f2' }, border: { bottom: { style: 'thin' } } })
  const total = st({ font: { bold: true } })
  const cells: Record<string, MemoryCell> = {
    A1: { value: 'Показатель', style: head },
    B1: { value: 'Значение', style: head },
    A2: { value: 'Цена' },
    B2: { value: 120 },
    A3: { value: 'Количество' },
    B3: { value: 450 },
    A4: { value: 'Выручка' },
    B4: { value: 54000, text: '54 000', formula: 'B2*B3' },
    A5: { value: 'Затраты на единицу' },
    B5: { value: 70 },
    A6: { value: 'Затраты' },
    B6: { value: 31500, text: '31 500', formula: 'B3*B5' },
    A7: { value: 'Прибыль до налога' },
    B7: { value: 22500, text: '22 500', formula: 'B4-B6' },
    A8: { value: 'Налог' },
    B8: { value: 4500, text: '4 500', formula: 'B7*E2' },
    A9: { value: 'Чистая прибыль', style: total },
    B9: { value: 18000, text: '18 000', formula: 'B7-B8', style: total },
    A10: { value: 'Рентабельность' },
    B10: { value: 1 / 3, text: '33,3%', numberFormat: '0.0%', formula: 'B9/B4' },
    A11: { value: 'План по прибыли' },
    B11: { value: 20000, text: '20 000', formula: 'Формулы!B2*100' },
    A12: { value: 'Выполнение плана', style: total },
    B12: { value: 0.9, text: '90,0%', numberFormat: '0.0%', formula: 'ЕСЛИ(B11>0;B9/B11;0)', style: total },
    D1: { value: 'Параметр', style: head },
    E1: { value: 'Значение', style: head },
    D2: { value: 'Ставка налога' },
    E2: { value: 0.2, text: '20%', numberFormat: '0%' },
  }
  return { name: 'Расчёт', cells, widths: { A: 22, B: 14, C: 3, D: 16, E: 10 } }
}

function merges(): MemorySheet {
  const center = st({ alignment: { horizontal: 'center', vertical: 'center' }, fill: { color: '#fff2cc' }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' } } })
  const right = st({ border: { top: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }, fill: { color: '#fff2cc' } })
  const cells: Record<string, MemoryCell> = {
    A1: { value: 'Шапка через пять колонок', style: center },
    E1: { style: right },
    A3: { value: 'Два на два', style: center },
    D3: { value: 'Справа' },
    A6: { value: 'Длинное объединение на 300 строк', style: st({ alignment: { vertical: 'top', wrap: true }, fill: { color: '#e2efda' } }) },
  }
  for (let r = 6; r <= 400; r++) cells[`C${r}`] = { value: r }
  return { name: 'Объединения', cells, merges: ['A1:E1', 'A3:B4', 'A6:B305'], widths: { A: 14, B: 14 } }
}

function sizes(): MemorySheet {
  const cells: Record<string, MemoryCell> = {}
  for (let c = 1; c <= 8; c++) for (let r = 1; r <= 12; r++) cells[cellAddress(r, c)] = { value: `${cellAddress(r, c)}` }
  cells.A1 = { value: 'B и строка 4 скрыты' }
  return {
    name: 'Размеры',
    cells,
    widths: { A: 22, B: 0, C: 4, D: 30, E: 9.140625 },
    heights: { '2': 30, '3': 8, '4': 0, '6': 45 },
    defaultRowHeight: 18,
  }
}

/**
 * Excel's outline: a budget's rows in two levels (items under their sections, the sections under the profit), the costs'
 * items collapsed; months under their quarters, the second quarter collapsed, and both quarters under the half-year.
 */
function grouping(): MemorySheet {
  const head = st({ font: { bold: true }, fill: { color: '#d9e1f2' }, border: { bottom: { style: 'thin' } } })
  const total = st({ font: { bold: true } })
  const sum = st({ font: { bold: true }, border: { top: { style: 'thin' } } })
  const titles = ['Статья', 'Янв', 'Фев', 'Мар', 'Кв. 1', 'Апр', 'Май', 'Июн', 'Кв. 2', 'Полугодие']
  const cells: Record<string, MemoryCell> = {}
  titles.forEach((t, i) => (cells[cellAddress(1, i + 1)] = { value: t, style: head }))
  const lines: [string, number[], number | undefined][] = [
    ['Продукт А', [120, 130, 125, 140, 150, 160], undefined],
    ['Продукт Б', [80, 85, 90, 95, 92, 99], undefined],
    ['Продукт В', [40, 42, 47, 51, 55, 58], undefined],
    ['Выручка', [], total],
    ['Сырьё', [-60, -62, -64, -70, -72, -75], undefined],
    ['Упаковка', [-12, -12, -13, -14, -14, -15], undefined],
    ['Логистика', [-20, -21, -22, -24, -25, -26], undefined],
    ['Зарплата', [-70, -70, -70, -72, -72, -72], undefined],
    ['Расходы', [], total],
    ['Прибыль', [], sum],
  ]
  const months = [2, 3, 4, 6, 7, 8]
  const value = (r: number, c: number): number => {
    const [, nums] = lines[r - 2]!
    if (nums.length) return nums[months.indexOf(c)]!
    if (r === 5) return value(2, c) + value(3, c) + value(4, c)
    if (r === 10) return value(6, c) + value(7, c) + value(8, c) + value(9, c)
    return value(5, c) + value(10, c)
  }
  lines.forEach(([name, , style], i) => {
    const r = i + 2
    cells[`A${r}`] = { value: name, style }
    for (const c of months) cells[cellAddress(r, c)] = { value: value(r, c), style }
    const q1 = value(r, 2) + value(r, 3) + value(r, 4)
    const q2 = value(r, 6) + value(r, 7) + value(r, 8)
    cells[cellAddress(r, 5)] = { value: q1, style: total }
    cells[cellAddress(r, 9)] = { value: q2, style: total }
    cells[cellAddress(r, 10)] = { value: q1 + q2, style: sum }
  })
  return {
    name: 'Группировка',
    cells,
    widths: { A: 16, 'F:H': 0 },
    heights: { '6:9': 0 },
    freeze: { rows: 1, cols: 1 },
    outline: { rows: { '2:4': 2, 5: 1, '6:9': 2, 10: 1 }, columns: { 'B:D': 2, E: 1, 'F:H': 2, I: 1 }, collapsed: ['10', 'I'] },
  }
}

function freeze(): MemorySheet {
  const cells: Record<string, MemoryCell> = {}
  const head = st({ font: { bold: true }, fill: { color: '#d9e1f2' }, border: { bottom: { style: 'thin' } } })
  for (let c = 1; c <= 30; c++) cells[cellAddress(1, c)] = { value: c === 1 ? 'Строка' : `Колонка ${c}`, style: head }
  for (let r = 2; r <= 300; r++) {
    cells[cellAddress(r, 1)] = { value: `Строка ${r}`, style: head }
    for (let c = 2; c <= 30; c++) cells[cellAddress(r, c)] = { value: r * 100 + c }
  }
  return { name: 'Закрепление', cells, freeze: { rows: 1, cols: 1 } }
}

function table(): MemorySheet {
  const cells: Record<string, MemoryCell> = { B2: { value: 'Месяц' }, C2: { value: 'Доход' }, D2: { value: 'Расход' } }
  const months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь']
  months.forEach((m, i) => {
    cells[`B${i + 3}`] = { value: m }
    cells[`C${i + 3}`] = { value: 1000 + i * 120 }
    cells[`D${i + 3}`] = { value: 700 + i * 90 }
  })
  return { name: 'Таблица', cells, tables: [{ name: 'Бюджет', range: 'B2:D8', headerRow: 2 }], widths: { B: 14, C: 12, D: 12 } }
}

function long(rows: number): MemorySheet {
  const cells: Record<string, MemoryCell> = { A1: { value: 'N' }, B1: { value: 'Код' }, C1: { value: 'Сумма' }, D1: { value: 'Признак' } }
  for (let r = 2; r <= rows; r++) {
    cells[`A${r}`] = { value: r - 1 }
    cells[`B${r}`] = { value: `K-${String(r).padStart(6, '0')}` }
    cells[`C${r}`] = { value: ((r * 7919) % 100000) / 100 }
    cells[`D${r}`] = { value: r % 3 === 0 }
  }
  return { name: 'Длинный лист', cells, freeze: { rows: 1, cols: 0 }, widths: { B: 12 } }
}

function locked(): MemorySheet {
  const cells: Record<string, MemoryCell> = { A1: { value: 'Скрипт' }, B1: { value: 'Скрипт' }, C1: { value: 'Пользователь' } }
  for (let r = 2; r <= 10; r++) {
    cells[`A${r}`] = { value: r }
    cells[`B${r}`] = { value: r * 2, formula: `A${r}*2` }
  }
  return { name: 'Запертые', cells, locked: [{ range: 'A1:B10', reason: 'это область скрипта' }], widths: { C: 14 } }
}

function overflow(): MemorySheet {
  return {
    name: 'Перелив',
    cells: {
      A1: { value: 'Длинный заголовок отчёта переливается на пустые ячейки справа', style: st({ font: { bold: true, size: 14 } }) },
      A3: { value: 'Текст влево до занятой ячейки' },
      D3: { value: 'стоп' },
      E5: { value: 'Текст вправо переливается налево', style: st({ alignment: { horizontal: 'right' } }) },
      C7: { value: 'По центру в обе стороны', style: st({ alignment: { horizontal: 'center' } }) },
      A9: { value: 123456789, text: '123456789' },
      A11: { value: 'Перенос не переливается', style: st({ alignment: { wrap: true } }) },
    },
  }
}

/** A table to filter, sort and search: a frozen header, repeating departments and statuses, sums and dates. */
function filterable(rows: number): MemorySheet {
  const head = st({ font: { bold: true }, fill: { color: '#d9e1f2' }, border: { bottom: { style: 'thin' } } })
  const cells: Record<string, MemoryCell> = {
    A1: { value: 'Отдел', style: head },
    B1: { value: 'Статус', style: head },
    C1: { value: 'Сумма', style: head },
    D1: { value: 'Дата', style: head },
    E1: { value: 'Комментарий', style: head },
  }
  const depts = ['Бухгалтерия', 'Склад', 'Продажи', 'Закупки', 'ИТ']
  const statuses = ['оплачен', 'ожидает', 'отклонён', '']
  for (let r = 2; r <= rows; r++) {
    const sum = ((r * 7919) % 50000) / 100
    const day = 1 + ((r * 13) % 28)
    cells[`A${r}`] = { value: depts[r % depts.length]! }
    const status = statuses[(r * 3) % statuses.length]!
    if (status) cells[`B${r}`] = { value: status }
    cells[`C${r}`] = { value: sum, text: sum.toFixed(2).replace('.', ',') }
    cells[`D${r}`] = { value: 46022 + day, text: `${String(day).padStart(2, '0')}.01.2026` }
    if (r % 7 === 0) cells[`E${r}`] = { value: `счёт ${r}` }
  }
  return { name: 'Фильтр', cells, freeze: { rows: 1, cols: 0 }, widths: { A: 16, B: 12, C: 12, D: 12, E: 18 } }
}

/** Excel's data validation: dropdown lists, input messages, whole numbers within bounds, and the three styles of alert. */
function validation(): MemorySheet {
  const head = st({ font: { bold: true }, fill: { color: '#d9e1f2' }, border: { bottom: { style: 'thin' } } })
  const months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
  const rows: [string, MemoryCell['value'], string][] = [
    ['Статус', 'открыт', 'список и подсказка'],
    ['Отдел', 'Склад', 'список, «Останов»'],
    ['Количество', 5, 'целое от 1 до 100, «Останов» со своим заголовком'],
    ['Скидка, %', 10, 'целое от 0 до 30, «Предупреждение»'],
    ['Приоритет', 'средний', 'список, «Сообщение»'],
    ['Оценка', 3, 'целое от 1 до 5, «Останов» без своего текста'],
    ['Месяц', 'Март', 'список из двенадцати пунктов'],
    ['Комментарий', null, 'только подсказка'],
  ]
  const cells: Record<string, MemoryCell> = { A1: { value: 'Поле', style: head }, B1: { value: 'Значение', style: head }, C1: { value: 'Правило', style: head } }
  rows.forEach(([name, value, rule], i) => {
    cells[`A${i + 2}`] = { value: name }
    if (value !== null) cells[`B${i + 2}`] = { value }
    cells[`C${i + 2}`] = { value: rule }
  })
  return {
    name: 'Проверка данных',
    cells,
    widths: { A: 14, B: 16, C: 46 },
    validations: [
      { range: 'B2', list: ['открыт', 'в работе', 'закрыт'], prompt: { title: 'Статус', text: 'Выберите статус из списка или нажмите Alt+↓' } },
      { range: 'B3', list: ['Бухгалтерия', 'Склад', 'Продажи', 'Закупки', 'ИТ'], alert: { message: 'Такого отдела нет: выберите из списка' } },
      { range: 'B4', whole: { min: 1, max: 100 }, prompt: { title: 'Количество', text: 'Целое число от 1 до 100' }, alert: { title: 'Количество', message: 'Введите целое число от 1 до 100' } },
      { range: 'B5', whole: { min: 0, max: 30 }, alert: { severity: 'warning', title: 'Скидка', message: 'Скидка больше 30 % требует согласования' } },
      { range: 'B6', list: ['низкий', 'средний', 'высокий'], alert: { severity: 'information', message: 'Обычно приоритет берут из списка' } },
      { range: 'B7', whole: { min: 1, max: 5 } },
      { range: 'B8', list: months },
      { range: 'B9', prompt: { title: 'Комментарий', text: 'Коротко, до 200 знаков: что сделано и что осталось' } },
    ],
  }
}

export function allFormattingBook(longRows = 34_000): MemoryBook {
  const sheets: MemorySheet[] = [
    formatting(),
    formats(),
    bars(),
    conditional(),
    links(),
    notes(),
    formulas(),
    chain(),
    merges(),
    sizes(),
    grouping(),
    { name: 'Скрытый', hidden: true, cells: { A1: { value: 'Скрытый лист: во вкладках его нет' } } },
    freeze(),
    table(),
    long(longRows),
    locked(),
    overflow(),
    { name: 'Плашка', cells: { A1: { value: 'Значения посчитал источник' } }, notice: 'Значения формул посчитаны источником, в файле их нет' },
    filterable(300),
    validation(),
    listSheet(),
  ]
  return { sheets, styles }
}
