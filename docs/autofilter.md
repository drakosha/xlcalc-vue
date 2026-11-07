# The AutoFilter

Excel's AutoFilter (Data > Filter) is a range of a sheet, outside a table, with a filter button on each cell of its header row. A source tells the sheet where a sheet's AutoFilter is with `SheetInfo.autoFilter` (`{ range: 'A1:D20', criteria? }`, the header row first), and turns it on or off with `setAutoFilter(sheet, range | null)`, a single undo step. `./local` gives the range the AutoFilter acts on ([below](#data-typed-under-it)) and sets it through `setLayout`.

## The buttons and the rows they filter

On a sheet with an AutoFilter, the display filter's funnels are Excel's buttons at the bottom right of the header row's cells, over the AutoFilter's columns only; the column letters have none. A button opens the same list as a funnel: the column's values over the AutoFilter's rows below the header, a text condition and the sort. The conditions and the sort apply to those rows only: the rows above the header row and the rows below the AutoFilter's last row stay where they are, so `visibleRows` gets `head` (the header row) and `last` (the last row) in its `SheetFilter`. A filtered column's button shows a funnel, a sorted one its direction, and the status line counts the AutoFilter's rows. Alt+Down on a header cell opens its button's list, as in Excel. A button is 16 px, up to 24 px on a touch screen, and never larger than its cell; with an AutoFilter the formula bar of a touch screen has no funnel of its own.

The ribbon's methods work on the AutoFilter's columns: `filterColumn()` opens the active cell's column's list, `sortColumn(dir)` sorts by it, and both return `false` for a column outside the AutoFilter. `clearFilter()` is Data > Clear and `reapplyFilter()` Data > Reapply, as for the display filter ([view.md](view.md#reapply)).

The data do not change: the display filter hides and orders rows on screen only, its sort as well, where Excel's rewrites the rows. With xlcalc 0.88, `SheetInfo.autoFilter.criteria` carries the saved conditions and sort. Value selections (including blanks), literal substring conditions and a single-column value sort appear in the menu and indicators. `./local` has `setFilterCriteria`: applying supported conditions, Clear and Reapply update the saved criteria and hidden rows in one undo step; Undo and Redo restore both. Sorting still orders the display only: its metadata is saved, but the underlying cells are not reordered.

Dates grouped by year or month, comparisons beyond a literal substring, Top 10, dynamic filters, colors and icons stay stored without being evaluated. Their columns are marked as filtered and the menu displays `filterStored`. While such conditions remain, applying another filter and Reapply are disabled or refused rather than guessing which rows to show. Reset can remove a stored condition, or Clear can remove all of them. Complex sorts stay stored, with their direction shown, but are not executed; choosing a simple sort replaces the saved sort. Button settings (`hiddenButton`, `showButton`) are respected and retained.

Without `setFilterCriteria`, including a read-only view, filtering remains a display choice and does not write to the workbook. Filters of Excel tables and standalone sheet sorts remain preserved by xlcalc; this menu manages the sheet's AutoFilter outside tables.

## Data > Filter

If the source has `setAutoFilter` and the sheet is editable, Ctrl+Shift+L (Control on a Mac too, with any keyboard layout) and the sheet's `toggleAutoFilter()` are Excel's Data > Filter. On a sheet with an AutoFilter it removes the AutoFilter with its conditions and sort; on a filtered sheet `./local` also shows the rows of its range below the header, those hidden by hand too, and marks the sheet as no longer filtered, as Excel's button does; other rows stay hidden, and a row of height 0 keeps its height. On a sheet without one it puts an AutoFilter over the range Excel picks, measured in Excel 16 for Mac on 29 cases:

- a single cell: its current region (the block of data around it, an empty cell next to the data included), trimmed to its rows and columns with data, so an empty cell just beside or under a table takes the table;
- a selection of more than one cell (entire rows and columns too): the selection trimmed to its cells with data, then taken down to the bottom of the data around it, so a selected header row takes the rows of data under it;
- no data there: nothing is sent, and the sheet displays Excel's message (`autoFilterNoRange`).

A new AutoFilter starts without conditions. `state.autoFilter` is the sheet's AutoFilter range, for a pressed Filter button on a ribbon, and `state.canAutoFilter` says Data > Filter is available. The source refuses an AutoFilter over a table, which has a filter of its own (`./local` with xlcalc's message), and a protected sheet may refuse it too; the sheet displays the refusal.

A source with `setAutoFilter` shows no funnels on a sheet without an AutoFilter, as Excel has none until Data > Filter. A source without it, or a read-only sheet, keeps the funnels on the column letters of such a sheet, and the buttons of a sheet that has one ([view.md](view.md#the-display-filter-on-a-ribbon)); list mode is unchanged.

## Data typed under it

Excel does not change an AutoFilter's range when data is typed under it, and saves the range as it was, but it filters more: the range it acts on goes down to the last row of its current region, so a row typed right under the data joins it, and a row past an empty one does not (xlcalc's `getAutoFilterRange`, measured in Excel for Mac and Windows). `SheetInfo.autoFilter` is that range: the buttons stay on the same header row over the same columns, and the display filter's conditions, its funnels' lists and the status line count the rows down to its last row. `./local` lists it, keeps the stored range in the workbook it saves, after typing, a paste, a fill, Delete or an undo that changes it sends `sheets` (a command after which it sends `reload` lists it anew), on which the sheet reads the sheets again and keeps its windows and the display filter's rows. As in Excel, the rows already filtered stay as they are: a row typed under a filtered range shows until Reapply (Ctrl+Alt+L) or a change of the conditions filters it with the others.

## When the sheet changes

The source moves the AutoFilter with rows and columns (`./local` through xlcalc: rows inserted inside it grow it, deleting its header row removes it, a table over all of it takes its place), and the sheet draws the buttons where the source puts them after it reloads. Columns inserted or deleted through the sheet's own commands move the conditions and the sort with their columns, and those of deleted columns go. An AutoFilter moved otherwise (an undo, a redo, another viewer's change) takes them along when it moved whole; when its width changed, the sheet cannot tell which columns moved, and conditions on columns outside it go. An AutoFilter that goes takes its conditions along.

`SheetToolbar` with `sheet` has this button, the group `autoFilter` after AutoSum (`auto-filter`, named `autoFilterToggle`): pressed while the sheet has an AutoFilter, disabled without `canAutoFilter`, and large in the two-row ribbon, as on Excel's Data tab. A host with a ribbon of its own draws it from the same state:

```ts
// A host's ribbon: the Filter button pressed while the sheet has an AutoFilter.
const filterButton = { id: 'filter', label: 'Filter', run: () => sheet.value?.toggleAutoFilter(), pressed: !!state.value?.autoFilter, disabled: !state.value?.canAutoFilter }
```
