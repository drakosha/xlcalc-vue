# PivotTables

A workbook's PivotTables show on the sheet as their cells, as Excel last laid them out. The sheet knows where each one lies, refreshes the one under the active cell, and shows Excel's alert for what a PivotTable's cells do not take. The PivotTable Fields pane, Insert PivotTable and the dialogs over a PivotTable are not part of the sheet yet; a host builds them over the source's `structure`, which takes xlcalc's operations on PivotTables ([Operations](#operations)).

## Where the PivotTables are

A source lists each sheet's PivotTables in `SheetInfo.pivots`, read with the sheets and again on `reload` or `sheets`. Each is a `SheetPivot`, its name and its areas as ranges:

| Field | Area |
| --- | --- |
| `name` | The PivotTable's name, unique among the sheet's (`PivotTable1`): an operation names the PivotTable by its sheet and this name. |
| `range` | The table without its Filters area (Excel's `TableRange1`). |
| `filters` | The Filters area above the table, as wide as its filters, without the blank row between (Excel's `PageRange`). |
| `headers` | The table's rows above its values: the value fields' and column fields' headers and the column labels. |
| `rowLabels` | The row labels left of the values, down to the grand total. |
| `values` | The values with their totals (Excel's `DataBodyRange`); absent without a value field. |

`@xlcalc/vue/local` takes them from xlcalc's `PivotLocation` (`range`, `firstHeaderRow`, `firstDataRow`, `firstDataColumn`, `pageRows`, `pageColumns`). A PivotTable not laid out (added without the engine: Excel's empty PivotTable of three columns and 18 rows) or shrunk to its top-left cell, which shows `#SPILL!`, is listed by its table (and Filters area) alone, without `headers`, `rowLabels` or `values`. A PivotTable laid out again (Refresh, a label renamed, an undo) is listed anew as the sheet reads the sheets again.

The sheet's `state` event names the PivotTable under the active cell, its Filters area included, as `pivot`, so that a ribbon can show Excel's PivotTable Analyze and Design tabs while the active cell is in one; `canRefreshPivot` says that Refresh is available there.

## Refresh and Refresh All

Excel's PivotTable Analyze > Refresh (Alt+F5) reads the PivotTable's source again and lays out every PivotTable over the same pivot cache. The sheet sends it to the source's `structure` as xlcalc's `refreshPivots` naming the PivotTable under the active cell: by Alt+F5 on the grid, by Refresh in the cell menu of a PivotTable's cell, by `SheetViewApi.refreshPivot()` for a ribbon, and by Refresh in the toolbar's Refresh All list. It is one step of undo, as xlcalc makes it; the sheet reads its windows again afterwards. A PivotTable keeps its values while its source changes, until a refresh, as in Excel.

Excel's Data > Refresh All (Ctrl+Alt+F5) refreshes the PivotTables too. It is the source's `refreshConnections` without `names`, so the sheet, `SheetViewApi.refreshAll()` and the toolbar send it as before ([calculation.md](calculation.md#refresh-all)). In `@xlcalc/vue/local` it asks the calls outside the workbook again first, when the host answers them (`externals.resolve`), then refreshes every pivot cache the workbook has then (xlcalc's `refreshPivots` without `pivots`, which passes over a cache over another workbook, a connection or the data model). The PivotTables are refreshed as soon as the calls return: an answer that arrives later (a cell that showed `#BUSY!` or `#GETTING_DATA`) reaches its cell, not the PivotTables over it, until the next refresh. Every `./local` source has Refresh All, and the workbook's layout at the call tells whether there are PivotTables to refresh; without them and without a resolver it changes nothing. Asking the calls again is not a step of undo, as in Excel; the PivotTables' refresh is one step, as xlcalc's. `refreshConnections` with `names` asks those functions' calls and leaves the PivotTables.

The toolbar's `refreshAll` group is a split button, as Excel's Refresh All on the Data tab: the button refreshes all, and its list holds Refresh All and Refresh, the latter enabled while the active cell is in a PivotTable. Its texts are `refreshAll`, `refreshAllMenu`, `refreshAllItem`, `menuRefresh` (the list's and the cell menu's Refresh) and `refreshPivot` (Refresh's tooltip with its key).

## What its cells take

xlcalc does not let an entry change a PivotTable's values, as Excel does not, and the sheet shows the source's refusal (`EditOutcome` code `pivot`) as an alert. `@xlcalc/vue/local` takes the kind of Excel's alert from xlcalc's refusal and words it as Excel for Windows does, in the workbook's language:

- a value typed, pasted or entered with Ctrl+Enter over the PivotTable, and a range copied or moved onto it: "We can't change this part of the PivotTable.";
- a formula typed over a label: "Cannot enter a formula for an item or field name in a PivotTable report.";
- Delete, Clear All, and inserting or deleting rows, columns or cells through it, moving part of it, filling or merging over it: "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again."

The English alerts are those Excel 16.0 for Windows answered when xlcalc measured these entries through its object model (its set `edit`); the Russian ones were not measured, nor was the alert for a range copied or moved onto a PivotTable. Every other refusal keeps xlcalc's words, which name Excel's reason: a value field named as a field of the source ("PivotTable field name already exists"), a Filters area's item it does not have ("is not an item of this field"), Refresh of a source that is gone. Text typed over a label renames it, as in Excel and as xlcalc does: an item's label gives the item a caption, a field's name, a value field's name, `Grand Total` or `Row Labels` their names, a Filters area's item shows that item.

## Protection

On a protected sheet, Use PivotTable and PivotChart (`SheetProtection.pivotTables`, "Use PivotTable reports" in Microsoft's Protect a worksheet) lets users "Format, change the layout, refresh, or otherwise modify PivotTable reports, or create new reports", and Edit objects (`objects`) lets them change objects not unlocked before the sheet was protected; slicers and timelines, the article says, "are already locked when you add them". `@xlcalc/vue/local` follows these words, not measured in Excel:

- Refresh, Refresh All of the PivotTables, and every operation on a PivotTable need Use PivotTable and PivotChart on each sheet whose PivotTables it lays out again: a refresh or a change of the cache (grouping, calculated fields and items) every PivotTable over the cache, a move both sheets.
- A slicer's or a timeline's choice and its Report Connections refilter PivotTables: the box on the sheets of the PivotTables it filters, and of those it connects.
- Adding or deleting a slicer or a timeline needs Edit objects on its sheet too.
- Text typed over a label renames it, a change of the PivotTable: besides an unlocked cell, it needs the box.

Without them the source refuses with Excel's protected sheet alert (`EditOutcome` code `protected`), and the cell menu's Refresh is disabled. Refresh All with a sheet that does not allow it asks the calls outside the workbook and refuses to refresh the PivotTables. The alert Excel shows for these was not measured, and Excel's own refusal may differ from Microsoft's description.

## Operations

The source's `structure` takes every operation of xlcalc on PivotTables, slicers and timelines (`addPivotTable`, `editPivotTable`, `renamePivotTable`, `movePivotTable`, `deletePivotTable`, `refreshPivots`, `groupPivotField`, `ungroupPivotField`, `renamePivotGroup`, `setPivotCalculatedField`, `setPivotCalculatedItem`, `addSlicer`, `editSlicer`, `deleteSlicer`, `addTimeline`, `editTimeline`, `deleteTimeline`), each one step of undo, so a host's PivotTable Fields pane or Insert PivotTable dialog sends them through `SheetViewApi.structure` and the sheet reads the sheets again afterwards. xlcalc's guide to PivotTables (its `docs/pivot-tables.md`) describes them. In `@xlcalc/vue/local` the engine is loaded when the workbook opens with PivotTables, otherwise before the first operation on PivotTables ([local.md](local.md#pivottables)), so a PivotTable added to a workbook without them is laid out at once. Where the engine cannot be loaded (its chunk fails to download), these operations, Refresh and the PivotTables' part of Refresh All refuse with the source's own message (`pivot`), "PivotTables can't be refreshed or changed now: this page could not load what they need. Reload the page and try again.", and the workbook stays as it is.

## Files

`openXlsx` reads a file's PivotTables, caches, slicers and timelines, and `saveXlsx` writes them; what the workbook does not hold of them (a part of a cache or a PivotTable, PivotTables the model cannot take, which xlcalc words as `pivotUnread`, and a cache's source on a deleted sheet, written as Excel writes it, `pivotSourceLinked`) is reported under the issue code `pivotTable`, which `groupIssues` groups as the others: one group, with the first issue's message. Writing over the original file (`over`) rejects slicers and timelines added or deleted with xlcalc's `KeptRefusal` `code: 'slicers'` ([local.md](local.md#files)).

## Not yet

The PivotTable's style is kept for Excel and not drawn: its cells show their own formats only. Slicers and timelines are not drawn, and the sheet has no field buttons on a PivotTable's headers.
