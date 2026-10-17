# List mode

A host's list, such as a view or a report from a database, goes on the same sheet as a book. The host's source gives one sheet of rows, and a few fields of the contract and props of `SheetView` turn the sheet into a list: titles in place of column letters, sorting by the host, drill-down cells, rows that open, and a totals row. The list keeps what the sheet has: selection, copying as Excel copies, selection totals, the keyboard.

## What the source gives

- `SheetInfo.columns`: a title for each column in place of its letter, its `align` (`left`, `center`, `right`) for the title and the totals, and `sortable` when a click on the title sorts.
- `SheetInfo.rowNumbers: false`: no row numbers on the left.
- `SheetInfo.freeze`: the header rows and the first columns kept in view, as in a book. Rows within `freeze.rows` hold the header and do not open.
- `SheetInfo.totals`: the totals row, the shown text by the sheet's columns (`null` for an empty cell). It stays at the bottom of the frame as frozen rows stay at its top, and the selection scrolls so the active cell is not under it.
- `SheetWindow.drill`: the cells that are drill-downs, by address. They are underlined; an empty cell is no drill-down.
- No `edit` for a list that is read only, and `formulaBar: false` when the list has no formulas to show.

## Sorting

The host sorts, as a list from a server sorts on the server. A click on a sortable title emits `sort` with the column (1-based) and the order: ascending first, descending on the column the `sort` prop says is ascending. The host changes its query, sets the `sort` prop, which puts the arrow on the title, and sends `reload` from the source, after which the sheet reads its windows again. A title that is not sortable selects its column, as a letter does.

A whole list over `memorySource`, which plays the host's query here: the page sorts its rows and writes them over the sheet with `setCells`, as a server's list would answer a new query and send `reload`.

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

// The host's sort: the rows in the order asked, written over the sheet
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

A click on a drill-down cell, or Enter on it, emits `cellDrill` with the sheet and the cell's address; the selection stays. With `openRows` the first column's cells are links that emit `rowOpen` with the sheet and the row, and Enter on any cell that is not a drill-down opens its row too. Without `openRows` Enter moves as in a sheet. A click with Shift, Control or Command extends or adds to the selection and opens nothing.

## Look

The drill-down cells and the first column's links take `--xs-link`, the totals row `--xs-head-bg` with a line of `--xs-freeze-line` above it. The texts of the titles' and links' tooltips are `SheetMessages`' `sortColumn` and `openRow`.
