# Slicers and timelines

A workbook's slicers and timelines over PivotTables lie over the cells at their places, as Excel draws them, and filter their PivotTables as Excel's do: a click on a slicer's item, Ctrl+click, Multi-Select, Clear Filter, a run of a timeline's periods. The sheet sends each change to the source's `structure` as one of xlcalc's operations (`editSlicer`, `editTimeline`, `addSlicer`, `addTimeline`, `deleteSlicer`, `deleteTimeline`), one step of undo, and reads the sheets again afterwards, so the PivotTables they filter show the new choice and the slicers their items with data anew. A refresh and an undo reload them too.

## What the source gives

A source lists each sheet's slicers in `SheetInfo.slicers` and its timelines in `SheetInfo.timelines`, read with the sheets ([api.md](api.md#windows-and-sheets)). Each is xlcalc's `Slicer` or `Timeline` with what its cache holds:

| Field | Meaning |
| --- | --- |
| `name` | Unique among the workbook's slicers and timelines: an operation names it. |
| `caption`, `showCaption` (`showHeader` of a timeline) | The header's text, and whether the header shows. |
| `place` | xlcalc's `DrawingPlace`: the cells of its top-left and bottom-right corners (`from`, `to`) with offsets into them in EMU (914,400 to an inch, 9,525 to a pixel). A slicer or timeline without one is not drawn. |
| `field`, `pivots`, `pivotCache` | The source field it filters, the PivotTables it filters (`{ sheet, name }`), their pivot cache's id. |
| `items` (a slicer) | Its items in the slicer's order (`SheetSlicerItem`): `item`, the index `editSlicer` names it by; `label`, as the slicer shows it; `selected: false` for an item not chosen; `noData: true` for an item no record of which passes the other filters. |
| `columns`, `rowHeight`, `sort`, `customListSort`, `noData`, `showMissing` (a slicer) | Its columns of buttons, their height in EMU, and its cache's sorting and display of items. |
| `level`, `selectionLevel`, `selection`, `bounds` (a timeline) | The level shown (months by default), the level the dates were chosen at, the dates chosen (`{ start, end }`, none for all periods), the dates it spans (whole years). |

`@xlcalc/vue/local` gives them from the workbook's layout, where xlcalc keeps each slicer cache's items with whether each is chosen and has data after every calculation. An item's label is its value in its field's number format, a date in the short date, a blank `(blank)` (`(пусто)` in ru-RU), as xlcalc's PivotTable engine writes an item's label; the library gives no caption of a slicer's item of its own. `SheetInfo.pivots` gives each PivotTable's pivot cache (`cache`) and the cache's fields (`fields`, `{ name, dates? }`, without calculated fields) for Insert Slicer and Report Connections.

## How they are drawn

Each slicer and timeline is a box in the grid's scroller, at its place, so it scrolls and zooms with the cells. With frozen panes each pane draws its part of the box, as Excel draws an object in each pane: the part in the frozen rows or columns stays while the rest scrolls under them. Each pane of a split window draws its own. A box whose corner lies on a row the display filter leaves out is not drawn.

The look is a plain neutral one, the same for every style: the workbook's `style` (`SlicerStyleLight1`, `TimeSlicerStyleLight1` and the others) is kept for Excel and not drawn yet. The CSS variables `--xs-slicer-*` set its colors ([api.md](api.md#css-variables)).

- A slicer: a frame; the header with its caption, Multi-Select and Clear Filter; a button for each item in `columns` columns, `rowHeight` high (Excel's 0.26 inch without one). A chosen item is filled, an item not chosen is plain, an item with no data has grey text and a dashed border (where the cache marks them), and the items scroll when they do not fit. With Hide items with no data, those items are not shown.
- A timeline: the header with its caption and Clear Filter; the selection's label (All Periods, `Mar 2025`, `Jan - Jun 2025`) and the level picker (YEARS, QUARTERS, MONTHS, DAYS); the periods of the level between its bounds under their years (under their months for days), the chosen ones filled, scrolled sideways, opening at the first chosen period.

## Choosing

- A click on a slicer's item chooses it alone. Ctrl+click adds it or takes it away; so does a click while Multi-Select is on (the header's button, Alt+S, or the menu's Multi-Select). Shift+click chooses the run of items from the last one clicked, added to the choice with Ctrl or Multi-Select. Taking away the last item chosen clears the filter (not measured in Excel). Excel does not save Multi-Select in the file, and the sheet keeps it while the page lives.
- Clear Filter (the header's button, Alt+C, or the menu's Clear Filter from "X") chooses every item, or takes a timeline's dates away (`selection: null`).
- A press on a timeline's period chooses it, a drag over periods chooses the run, Shift+click extends the run from the last period pressed; the dates sent are the first day of the first period to the last day of the last, which xlcalc applies by whole days.
- The level picker changes the level shown (`editTimeline` `set.level`), which does not change the dates.

Each change is one step of undo, as xlcalc makes it. Without xlcalc's PivotTable engine the choice is written all the same and the PivotTables keep their cells until a refresh ([pivot-tables.md](pivot-tables.md#operations)).

## Selecting, moving, resizing, removing

A press on a slicer or a timeline selects it, as Excel's does: `state.object` names it, its frame takes the keys and shows its handles. A drag of the frame (its header, its edge) moves it, a drag of a handle resizes it; the release sends its new place (`set.place`, the corners' cells and offsets), as one step of undo. Delete removes it (`deleteSlicer`, `deleteTimeline`); Esc gives the keys back to the grid; a press on a cell leaves it.

A right click opens its menu, with the items of Excel's menu the sheet has: Refresh (the PivotTables it filters), Sort A to Z and Sort Z to A (a slicer), Clear Filter from "X", Multi-Select (a slicer), Remove "X", Report Connections…, Slicer Settings… (a slicer). Excel's Cut, Copy, Paste Options, Bring to Front, Send to Back, Size and Properties…, Assign Macro… and Format Slicer… are not there. The order follows Excel's string tables and has not been measured.

## Insert Slicer and Insert Timeline

`SheetViewApi.insertSlicer()` opens Excel's Insert Slicers dialog over the PivotTable under the active cell, a box for each field of its cache; OK adds a slicer of each field checked (`addSlicer` with the field and that PivotTable) and selects the last. `insertTimeline()` opens Insert Timelines with the fields of dates, or, without one, Excel's alert "We can't create a Timeline for this report because it doesn't have a field formatted as Date."; OK adds a timeline of each field checked. `state.canInsertSlicer` says they are available. xlcalc names a new slicer after its field (`Region`, `Region 1`) and places it; the sheet sends no place of its own.

## Slicer Settings and Report Connections

`slicerSettings()` (and the menu's Slicer Settings…) opens Excel's dialog for the selected slicer: Source Name and Name to use in formulas (the field and the slicer cache's name), Name (shown, not changed: xlcalc's `editSlicer` does not rename), Header with Display header and Caption, and Item Sorting and Filtering with Ascending (A to Z), Descending (Z to A), Use Custom Lists when sorting, Hide items with no data, Visually indicate items with no data, Show items with no data last and Show items deleted from the data source. OK sends them as one `editSlicer` (`set.caption`, `set.showCaption`, `cache.sort`, `cache.customListSort`, `cache.noData`, `cache.showMissing`). Excel's variants of the sort options for numbers and dates (Smallest to Largest, Oldest to Newest) are not drawn.

`reportConnections()` (and the menu's Report Connections…) opens Excel's Report Connections for the selected slicer or timeline: every PivotTable of the workbook over the same pivot cache, by name and sheet, those it filters checked. OK sends the PivotTables checked as `cache.pivots`; a PivotTable connected takes the slicer's choice, one disconnected shows every item again, as xlcalc does.

## Protection

Microsoft's Protect a worksheet says slicers and timelines are locked when added, that the lock works once the sheet is protected, and that Use PivotTable reports lets users "format, change the layout, refresh, or otherwise modify PivotTable reports"; its object model says a slicer whose `Locked` is false "can be modified when the sheet is protected". `@xlcalc/vue/local` checks a protected sheet ([pivot-tables.md](pivot-tables.md#protection)): a choice, a level, a place, settings and Report Connections need Use PivotTable and PivotChart on the sheets of the PivotTables the cache filters; adding and removing need Edit objects and Use PivotTable and PivotChart on its sheet. The sheet disables what the source refuses, and besides disables moving and resizing a locked slicer, or any timeline, without Edit objects, as Excel treats a locked object. None of this was measured in Excel; the alert is the protected sheet's.

## Texts

The texts are `SlicerMessages` of `SheetMessages`, in English and Russian. They are Excel's own words from its string tables and Microsoft's pages on slicers and timelines, and have not yet been compared with Excel for Windows on screen.
