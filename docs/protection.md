# Sheet and cell protection

Excel's Protect Sheet keeps users from changing a sheet except where it allows them to: they may type only into the cells that are not locked, and do only what the sheet's protection still allows (the boxes of Excel's Protect Sheet dialog). The workbook holds both, as xlcalc reads them from an xlsx file and writes them back: a sheet's protection is `SheetInfo.protection` (xlcalc's `SheetProtection`), and a cell is locked, or its formula hidden, by its style (`CellStyle.protection`, Excel's Format Cells, Protection). Every cell is locked unless its style unlocks it, so protecting a sheet without unlocking any cell keeps the whole sheet from being changed.

The workbook holds no password, so neither does the sheet: Protect Sheet has no password field, and Unprotect Sheet asks for none. A sheet protected by a password in an xlsx file is protected without it once read, and is saved without one (`writeKept` keeps the file's password while the protection is unchanged, xlcalc's docs/xlsx.md). Protection is no security either: in Excel and here the data are all in the workbook, and a hidden formula is only not displayed.

## The source

A source tells the sheet that a sheet is protected through `SheetInfo.protection`, with what it allows: `{}` is Excel's default, selecting cells and nothing else, and `{ formatCells: true }` also allows formatting cells. Each window's `styles` carry the cells' `protection`. `protect(sheet, protection)` protects a sheet (null unprotects it) as a single undo step; with it the sheet tab's menu offers Protect Sheet and Unprotect Sheet.

The source enforces the protection: the sheet only disables what Excel disables and checks the cells it has loaded. A change a protected sheet does not allow is refused with `reason: 'refused'`, `code: 'protected'` and Excel's alert as the message, which the sheet displays in an alert as Excel does.

## `./local`

xlcalc's session does not enforce protection, so `localSource` checks each change to a protected sheet before it reaches the session, and refuses it with Excel's own alert in the workbook's language, as Excel for Windows and Excel for Mac word it: "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password." (ru-RU: "Ячейка или диаграмма, которую вы пытаетесь изменить, находится на защищенном листе. Чтобы внести изменения, снимите защиту листа. Может потребоваться пароль."). What it allows was measured in Excel for Mac 16.114 through its object model, on a sheet with locked and unlocked cells protected with each box:

| Change | Allowed on a protected sheet |
| --- | --- |
| Typing, Ctrl+Enter, a paste, a dropdown's item, Delete, Replace All | when every cell it writes is unlocked |
| Clear Formats, Clear All | over unlocked cells; Clear Formats locks them again, and Clear All without Format cells clears their contents and keeps their formats |
| Formatting (`format`, conditional formatting rules) | with Format cells, over locked cells too; Locked and Hidden never change, nor does a cell's own style go |
| Column widths, hiding columns; row heights, hiding rows | with Format columns; with Format rows |
| Inserting rows, columns | with Insert rows, Insert columns |
| Deleting rows, columns | with Delete rows, Delete columns, when none of their cells is locked, refused otherwise with the same alert, as Excel for Windows shows it there |
| Pasting, moving and filling ranges (`copyRange`, `moveRange`, `fillRange`) | when every cell they write (and, for a move, take) is unlocked; Paste Special Formats with Format cells too, Column Widths with Format columns, Notes with Edit objects, Validation never |
| Hyperlinks | with Insert hyperlinks, over unlocked cells |
| Notes | with Edit objects, on locked cells too |
| Refresh, Refresh All's PivotTables, xlcalc's operations on PivotTables, slicers and timelines (`structure`), a label renamed by typing | with Use PivotTable and PivotChart on the sheet of each PivotTable it lays out again or filters, and Edit objects too for adding or deleting a slicer or a timeline, as Microsoft's Protect a worksheet describes the boxes (not measured in Excel); a refresh lays out every PivotTable over its cache ([pivot-tables.md](pivot-tables.md#protection)) |
| Merging and unmerging, the outline, tables, data validation rules | never; a table's refusal is Excel's "Table features aren't available because the sheet is protected." |
| Freeze Panes, sheets (insert, delete, rename, move, hide), defined names | always: they are the view and the workbook's structure, which a sheet's protection does not cover |

A cell beyond the sheet's used range is locked unless its row's style, or else its column's, unlocks it. Sort and AutoFilter belong to the AutoFilter, which the sheet's display filter does not change: it filters and sorts the display, not the data, on a protected sheet too.

## On the sheet

The sheet tab's menu has Protect Sheet… with the source's `protect`, or Unprotect Sheet on a protected sheet, as do the sheet's `protectSheet()` and `unprotectSheet()`. Protect Sheet is Excel's dialog without its password field: Protect worksheet and contents of locked cells (OK is disabled without it, as in Excel), then Allow all users of this worksheet to: with the boxes of what users may still do, in Excel's order and with its defaults (selecting cells only); Excel for Mac's note that all cells are locked by default is left out, as Excel for Windows has none. Its labels are Excel for Windows' in English and in Russian (`Защитить лист и содержимое защищаемых ячеек`, `Выделение заблокированных ячеек`); Unprotect Sheet is named as in both. Checking Select locked cells checks Select unlocked cells too, and unchecking Select unlocked cells unchecks Select locked cells, as Excel ties them. Unprotect Sheet asks nothing. Both are a single undo step, after which the sheet reads its sheets again.

Format Cells has Excel's Protection tab, Locked and Hidden with its note, in Excel for Windows' words (`Защищаемая ячейка`, `Скрыть формулы`) ([toolbar.md](toolbar.md#formatdialog)); on a protected sheet both boxes are disabled, as Excel changes neither there.

On a protected sheet, as in Excel:

- Typing into a locked cell, F2, Delete, a paste, Ctrl+Enter, the fill handle, Fill Down and Fill Right, Clear, Spelling's Change, a dropdown's item and Insert Function or Symbol show Excel's alert ("The cell or chart you're trying to change is on a protected sheet…", `protectedAlert`) before anything opens or is sent, for the cells the sheet has loaded; the source refuses the rest, and the sheet shows the same alert for its refusal (`code: 'protected'`).
- A hidden cell's formula is not displayed: the formula bar shows nothing and the `state` event has no `formula`; the cell shows its value, and its tooltip the value only. Show Formulas draws the cell empty, as Excel for Mac does (measured in 16.114).
- Trace Precedents (`explain()` and the cell menu's item) and Trace Dependents (`traceDependents()`) are disabled, as Excel disables formula auditing on a protected sheet.
- Formatting (the toolbar, Ctrl+B and the other shortcuts, Format Cells, Format Painter, Conditional Formatting) is disabled without Format cells. The row and column menus disable Insert and Delete without their boxes, Row Height, Column Width, Hide and Unhide without Format rows or Format columns, and Group and Ungroup always; the edges of row numbers and column letters do not resize without them either. Merge & Center, Data Validation, Format as Table and the outline's buttons are disabled; Insert Link without Insert hyperlinks, New Note, Edit Note and Show/Hide Note without Edit objects (with it they work on a locked cell too).
- Paste Special checks what the source checks before it sends: Notes need Edit objects and Column widths Format columns, over locked cells too, and Validation is refused; the other kinds need unlocked cells.
- Without Select locked cells, only unlocked cells can be selected: a click on a locked cell does nothing, an arrow goes on to the next unlocked cell that way (or stays where there is none), and Tab and Enter move to the next unlocked cell along the rows or down the columns, wrapping around, among the cells the sheet has loaded. Ctrl+arrow stays put when the cell it reaches is locked; a click on a column letter, a row number or Select All, and Ctrl+A, select nothing; Go To closes without moving when the range has a locked cell. The Name Box still goes to a locked cell, as Excel for Mac's does (measured in 16.114). Go To Special is disabled on a protected sheet, as in Excel for Mac.

The display filter and its sort change no data, so a protected sheet does not limit them, the AutoFilter's buttons included. Data > Filter (`toggleAutoFilter()`, Ctrl+Shift+L) is disabled on a protected sheet whatever its boxes allow, and `./local` refuses `setAutoFilter` there: in Excel the Sort and Use AutoFilter boxes let users use an AutoFilter the sheet has, not set or remove one ([autofilter.md](autofilter.md)).

Not as in Excel: Find's Look in Formulas finds text in a hidden formula, and Trace Precedents from a cell of another sheet explains a hidden formula among its precedents, as the source's `find` and `explain` answer; the sheet checks only the loaded cells before a change, so a change beyond them reaches the source, whose refusal it shows.
