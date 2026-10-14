// The example's list: a host's view in list mode, sorted by the page through the source, with drill-down sums, rows that open and totals.
import { cellAddress } from 'xlcalc/a1'
import type { ListSort, MemoryCell, MemorySheet, MemorySource } from '../src'

export const LIST_SHEET = 'Список'

interface Client {
  name: string
  region: string
  sum: number
  deals: number
}
const REGIONS = ['Москва', 'Казань', 'Пермь', 'Тверь', 'Омск']
const CLIENTS: Client[] = Array.from({ length: 200 }, (_, i) => ({
  name: `Клиент ${String(i + 1).padStart(3, '0')}`,
  region: REGIONS[(i * 7) % REGIONS.length]!,
  sum: ((i + 1) * 7919) % 100_000 / 100,
  deals: ((i + 1) * 37) % 23 + 1,
}))
const KEYS: (keyof Client)[] = ['name', 'region', 'sum', 'deals']

export function listSheet(): MemorySheet {
  const cells: Record<string, MemoryCell> = {}
  CLIENTS.forEach((client, i) => KEYS.forEach((k, c) => (cells[cellAddress(i + 1, c + 1)] = { value: client[k], ...(k === 'sum' ? { drill: true } : {}) })))
  const total = (k: 'sum' | 'deals') => CLIENTS.reduce((n, client) => n + client[k], 0)
  return {
    name: LIST_SHEET,
    rowNumbers: false,
    freeze: { rows: 0, cols: 1 },
    columns: [{ title: 'Клиент', sortable: true }, { title: 'Регион', sortable: true }, { title: 'Сумма', align: 'right', sortable: true }, { title: 'Сделок', align: 'right' }],
    totals: ['Итого', null, total('sum').toFixed(2), String(total('deals'))],
    cells,
    widths: { A: 16, B: 14, C: 12, D: 10 },
  }
}

/** The host's sort: the clients in the order asked, written over the sheet as another tab would write them. */
export function sortList(source: MemorySource, sort: ListSort): void {
  const k = KEYS[sort.col - 1]!
  const sorted = CLIENTS.slice().sort((a, b) => (a[k] < b[k] ? -1 : a[k] > b[k] ? 1 : 0) * (sort.dir === 'asc' ? 1 : -1))
  const entries: Record<string, string> = {}
  sorted.forEach((client, i) => KEYS.forEach((key, c) => (entries[cellAddress(i + 1, c + 1)] = String(client[key]))))
  source.setCells(LIST_SHEET, entries)
}
