// The local source's tests share these: a small ru-RU book, sources closed after each test, a window and a cell's style.
import { createInlineCalc, type AsyncSession, type BookInput, type Calc } from 'xlcalc'
import { localSource, type LocalSource, type LocalSourceOptions } from '../src/local'
import type { SheetFilter, SheetWindow } from '../src/contract/source'

export const BOLD = { font: { bold: true } }
export const ROW_FILL = { fill: { color: '#ffeeaa' } }
export const COL_FILL = { fill: { color: '#ddeeff' } }

/** A small book in ru-RU: values, a sum, styles of cells, rows and columns, a note, a link, a table, a list to filter. */
export const book = (): BookInput => ({
  locale: 'ru-RU',
  styles: [{}, BOLD, ROW_FILL, COL_FILL],
  sheets: [
    {
      name: 'Лист1',
      cells: {
        A1: { value: 1 },
        A2: { value: 2 },
        A3: { formula: 'SUM(A1:A2)' },
        B1: { value: 'с примечанием', note: { author: 'Аня', text: 'Проверить' } },
        B2: { value: 'ссылка', link: { url: 'https://example.org/', tooltip: 'Пример' } },
        B3: { value: 'на лист', link: { location: "'Таблица'!A1" } },
        B4: { value: 'место', link: { url: 'https://example.org/page', location: 'top' } },
        C1: { value: 'жирный', style: 1 },
        D5: { value: 'строка' },
      },
      columns: { B: { width: 20 }, D: { style: 3 }, E: { hidden: true } },
      rows: { '2': { height: 30 }, '5': { style: 2 } },
      merges: ['F1:G2'],
      frozen: { rows: 1, columns: 1 },
    },
    {
      name: 'Таблица',
      cells: {
        A1: { value: 'Товар' },
        B1: { value: 'Сумма' },
        C1: { value: 'Двойная' },
        A2: { value: 'чай' },
        B2: { value: 10 },
        C2: { formula: 'Продажи[[#This Row],[Сумма]]*2' },
        A3: { value: 'кофе' },
        B3: { value: 20 },
        C3: { formula: 'Продажи[[#This Row],[Сумма]]*2' },
        E1: { formula: 'SUM(Продажи[Сумма])' },
      },
      tables: [{ name: 'Продажи', range: 'A1:C3', columns: [{ name: 'Товар' }, { name: 'Сумма' }, { name: 'Двойная', formula: 'Продажи[[#This Row],[Сумма]]*2' }] }],
    },
    {
      name: 'Список',
      cells: {
        A1: { value: 'Имя' },
        B1: { value: 'Число' },
        A2: { value: 'б' },
        B2: { value: 5 },
        A3: { value: 'а' },
        B3: { value: 30 },
        A4: { value: 'б' },
        B4: { value: 'нет' },
        B5: { value: 7 },
        A6: { value: 'в' },
        B6: { value: 1.5 },
      },
    },
  ],
})

let open: { close(): Promise<void> }[] = []
/** A source over `b` computed in the same thread; `closeAll` closes it after the test. */
export function source(b: BookInput = book(), opts: LocalSourceOptions = {}): LocalSource {
  const s = localSource(b, { calc: 'inline', ...opts })
  open.push(s)
  return s
}
export async function closeAll(): Promise<void> {
  await Promise.all(open.map((s) => s.close()))
  open = []
}

/** A source whose session tells each call it gets (the method and its arguments), to see what a query reads from the book. */
export function watched(b: BookInput = book()): { s: LocalSource; calls: { method: string; args: unknown[] }[]; count: (method: string) => number } {
  const calls: { method: string; args: unknown[] }[] = []
  const inline = createInlineCalc()
  const calc: Calc = {
    compute: (...a) => inline.compute(...a),
    close: () => inline.close(),
    createSession: async (...a) => {
      const real = await inline.createSession(...a)
      return new Proxy(real, {
        get(target, key) {
          const v: unknown = Reflect.get(target, key, target)
          if (typeof v !== 'function') return v
          return (...args: unknown[]) => {
            calls.push({ method: String(key), args })
            return (v as (...x: unknown[]) => unknown).apply(target, args)
          }
        },
      }) as AsyncSession
    },
  }
  const s = source(b, { calc })
  open.push(calc)
  return { s, calls, count: (method) => calls.filter((c) => c.method === method).length }
}

export async function win(s: LocalSource, sheet: string, row: number, rows: number, col: number, cols: number, filter?: SheetFilter): Promise<SheetWindow> {
  const w = await s.window({ sheet, row, rows, col, cols, ...(filter ? { filter } : {}) })
  if ('pending' in w) throw new Error('pending')
  return w
}
export const styleOf = (w: SheetWindow, addr: string) => (w.styleIds[addr] === undefined ? undefined : w.styles[w.styleIds[addr]!])
