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

```vue
<!-- onSort: the host's own, it sets `sort`, changes the list's query and makes the source send `reload` -->
<SheetView :source="source" :sort="sort" open-rows read-only :formula-bar="false" @sort="onSort" @cell-drill="onDrill" @row-open="onRowOpen" />
```

## Drill-down and opening rows

A click on a drill-down cell, or Enter on it, emits `cellDrill` with the sheet and the cell's address; the selection stays. With `openRows` the first column's cells are links that emit `rowOpen` with the sheet and the row, and Enter on any cell that is not a drill-down opens its row too. Without `openRows` Enter moves as in a sheet. A click with Shift, Control or Command extends or adds to the selection and opens nothing.

## Look

The drill-down cells and the first column's links take `--xs-link`, the totals row `--xs-head-bg` with a line of `--xs-freeze-line` above it. The texts of the titles' and links' tooltips are `SheetMessages`' `sortColumn` and `openRow`.
