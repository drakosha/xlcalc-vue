# List mode

A host's list, such as a view or a report from a database, is displayed by the same sheet component as a workbook. The host's source returns a single sheet of rows, and a few fields of the contract and props of `SheetView` turn the sheet into a list: column titles in place of column letters, sorting by the host, drill-down cells, rows that open, and a totals row. The list keeps the sheet's features: selection, copying as Excel copies, selection totals, keyboard navigation.

## What the source returns

- `SheetInfo.columns`: a title for each column in place of its letter, its `align` (`left`, `center`, `right`) for the title and the totals, and `sortable` when clicking the title sorts.
- `SheetInfo.rowNumbers: false`: no row numbers on the left.
- `SheetInfo.freeze`: the header rows and the first columns kept in view, as in a workbook. Rows within `freeze.rows` are the header and do not open.
- `SheetInfo.totals`: the totals row, the displayed text for each of the sheet's columns (`null` for an empty cell). It stays at the bottom of the viewport as frozen rows stay at its top, and the sheet scrolls so the active cell is not hidden under it.
- `SheetWindow.drill`: the drill-down cells, by address. They are underlined; an empty cell is never a drill-down.
- No `edit` for a read-only list, and `formulaBar: false` when the list has no formulas to show.

## Sorting

The host sorts, as a server-side list is sorted on the server. Clicking a sortable title emits `sort` with the column (1-based) and the direction: ascending first, descending on the column the `sort` prop already shows as ascending. The host changes its query, sets the `sort` prop, which shows the arrow on the title, and sends `reload` from the source, after which the sheet reads its windows again. Clicking a title that is not sortable selects its column, as a column letter does.

A complete list over `memorySource`, which stands in for the host's query here: the page sorts its rows and writes them over the sheet with `setCells`, where a server-side list would answer a new query and send `reload`.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { SheetView, memorySource, type ListSort, type MemoryCell } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

interface Client { name: string; region: string; sum: number }
const clients: Client[] = Array.from({ length: 50 }, (_, i) => ({
  name: `Client ${String(i + 1).padStart(2, '0')}`,
  region: ['North', 'South', 'East', 'West'][i % 4]!,
  sum: ((i + 1) * 7919) % 10_000 / 10,
}))
const keys = ['name', 'region', 'sum'] as const
const letters = ['A', 'B', 'C']

// The list as cells: row i + 1 is clients[i]; the sums are drill-downs
function cellsOf(list: Client[]): Record<string, MemoryCell> {
  const cells: Record<string, MemoryCell> = {}
  list.forEach((client, i) => keys.forEach((k, c) => (cells[`${letters[c]}${i + 1}`] = { value: client[k], ...(k === 'sum' ? { drill: true } : {}) })))
  return cells
}

const source = memorySource({
  sheets: [
    {
      name: 'Clients',
      rowNumbers: false,
      freeze: { rows: 0, cols: 1 },
      columns: [{ title: 'Client', sortable: true }, { title: 'Region', sortable: true }, { title: 'Sum', align: 'right', sortable: true }],
      totals: ['Total', null, clients.reduce((n, c) => n + c.sum, 0).toFixed(1)],
      widths: { A: 14, B: 12, C: 12 },
      cells: cellsOf(clients),
    },
  ],
})

const sort = ref<ListSort | null>(null)
const opened = ref('')

// The host's sort: the rows in the requested order, written over the sheet
function onSort(next: ListSort): void {
  const k = keys[next.col - 1]!
  const sign = next.dir === 'asc' ? 1 : -1
  const sorted = clients.slice().sort((a, b) => (a[k] < b[k] ? -sign : a[k] > b[k] ? sign : 0))
  const entries: Record<string, string> = {}
  for (const [cell, { value }] of Object.entries(cellsOf(sorted))) entries[cell] = String(value)
  source.setCells('Clients', entries)
  sort.value = next
}
</script>

<template>
  <p>{{ opened }}</p>
  <SheetView
    :source="source" :sort="sort" open-rows read-only :formula-bar="false" style="height: 360px"
    @sort="onSort" @cell-drill="opened = `The deals behind ${$event.cell}`" @row-open="opened = `Client of row ${$event.row}`"
  />
</template>
```

## Drill-down and opening rows

Clicking a drill-down cell, or pressing Enter on it, emits `cellDrill` with the sheet and the cell's address; the selection stays. With `openRows`, the cells of the first column are links that emit `rowOpen` with the sheet and the row, and Enter on any cell that is not a drill-down also opens its row. Without `openRows`, Enter moves the active cell as on a regular sheet. A click with Shift, Ctrl or Cmd extends or adds to the selection and opens nothing.

## Appearance

Drill-down cells and the first column's links use `--xs-link`, and the totals row uses `--xs-head-bg` with a `--xs-freeze-line` line above it. The tooltips of the titles and links are `SheetMessages`' `sortColumn` and `openRow`.
