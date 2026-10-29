# Sheet and cell protection

Excel's Protect Sheet keeps users from changing a sheet except where it allows them to: they may type only into the cells that are not locked, and do only what the sheet's protection still allows (the boxes of Excel's Protect Sheet dialog). The workbook holds both, as xlcalc reads them from an xlsx file and writes them back: a sheet's protection is `SheetInfo.protection` (xlcalc's `SheetProtection`), and a cell is locked, or its formula hidden, by its style (`CellStyle.protection`, Excel's Format Cells, Protection). Every cell is locked unless its style unlocks it, so protecting a sheet without unlocking any cell keeps the whole sheet from being changed.

The workbook holds no password, so neither does the sheet: Protect Sheet has no password field, and Unprotect Sheet asks for none. A sheet protected by a password in an xlsx file is protected without it once read, and is saved without one (`writeKept` keeps the file's password while the protection is unchanged, xlcalc's docs/xlsx.md). Protection is no security either: in Excel and here the data are all in the workbook, and a hidden formula is only not displayed.

## The source

A source tells the sheet that a sheet is protected through `SheetInfo.protection`, with what it allows: `{}` is Excel's default, selecting cells and nothing else, and `{ formatCells: true }` also allows formatting cells. Each window's `styles` carry the cells' `protection`. `protect(sheet, protection)` protects a sheet (null unprotects it) as a single undo step; with it the sheet tab's menu offers Protect Sheet and Unprotect Sheet.

The source enforces the protection: the sheet only disables what Excel disables and checks the cells it has loaded. A change a protected sheet does not allow is refused with `reason: 'refused'`, `code: 'protected'` and Excel's alert as the message, which the sheet displays in an alert as Excel does.

## `./local`

xlcalc's session does not enforce protection, so `localSource` checks each change to a protected sheet before it reaches the session, and refuses it with Excel for Mac's own alert in the workbook's language: "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password." (ru-RU: "Ячейка или диаграмма, которую вы пытаетесь изменить, находится на защищенном листе. Чтобы внести изменения, снимите защиту листа. Может потребоваться пароль."). What it allows was measured in Excel for Mac 16.114 through its object model, on a sheet with locked and unlocked cells protected with each box:

| Change | Allowed on a protected sheet |
| --- | --- |
| Typing, Ctrl+Enter, a paste, a dropdown's item, Delete, Replace All | when every cell it writes is unlocked |
| Clear Formats, Clear All | over unlocked cells; Clear Formats locks them again, and Clear All without Format cells clears their contents and keeps their formats |
| Formatting (`format`, conditional formatting rules) | with Format cells, over locked cells too; Locked and Hidden never change, nor does a cell's own style go |
| Column widths, hiding columns; row heights, hiding rows | with Format columns; with Format rows |
| Inserting rows, columns | with Insert rows, Insert columns |
| Deleting rows, columns | with Delete rows, Delete columns, when none of their cells is locked, refused otherwise with Excel's "You are trying to delete a row that contains a locked cell…" |
| Pasting, moving and filling ranges (`copyRange`, `moveRange`, `fillRange`) | when every cell they write (and, for a move, take) is unlocked; Paste Special Formats with Format cells too, Column Widths with Format columns, Notes with Edit objects, Validation never |
| Hyperlinks | with Insert hyperlinks, over unlocked cells |
| Notes | with Edit objects, on locked cells too |
| Merging and unmerging, the outline, tables, data validation rules | never; a table's refusal is Excel's "Table features aren't available because the sheet is protected." |
| Freeze Panes, sheets (insert, delete, rename, move, hide), defined names | always: they are the view and the workbook's structure, which a sheet's protection does not cover |

A cell beyond the sheet's used range is locked unless its row's style, or else its column's, unlocks it. Sort and AutoFilter belong to the AutoFilter, which the sheet's display filter does not change: it filters and sorts the display, not the data, on a protected sheet too.

## On the sheet

The sheet tab's menu has Protect Sheet… with the source's `protect`, or Unprotect Sheet on a protected sheet, as do the sheet's `protectSheet()` and `unprotectSheet()`. Protect Sheet is Excel for Mac's dialog without its password field: Protect the sheet and contents of locked cells (OK is disabled without it, as in Excel), then the boxes of what users may still do, in Excel's order and with its defaults (selecting cells only), and Excel's note that all cells are locked by default. Checking Select locked cells checks Select unlocked cells too, and unchecking Select unlocked cells unchecks Select locked cells, as Excel ties them. Unprotect Sheet asks nothing. Both are a single undo step, after which the sheet reads its sheets again.

Format Cells has Excel's Protection tab, Locked and Hidden with its note ([toolbar.md](toolbar.md#formatdialog)); on a protected sheet both boxes are disabled, as Excel changes neither there.

On a protected sheet, as in Excel:

- Typing into a locked cell, F2, Delete, a paste, Ctrl+Enter, the fill handle, Fill Down and Fill Right, Clear, Spelling's Change, a dropdown's item and Insert Function or Symbol show Excel's alert ("The cell or chart you're trying to change is on a protected sheet…", `protectedAlert`) before anything opens or is sent, for the cells the sheet has loaded; the source refuses the rest, and the sheet shows the same alert for its refusal (`code: 'protected'`).
- A hidden cell's formula is not displayed: the formula bar shows nothing and the `state` event has no `formula`; the cell shows its value.
- Formatting (the toolbar, Ctrl+B and the other shortcuts, Format Cells, Format Painter, Conditional Formatting) is disabled without Format cells. The row and column menus disable Insert and Delete without their boxes, Row Height, Column Width, Hide and Unhide without Format rows or Format columns, and Group and Ungroup always; the edges of row numbers and column letters do not resize without them either. Merge & Center, Data Validation, Format as Table and the outline's buttons are disabled; Insert Link without Insert hyperlinks, New Note and Edit Note without Edit objects.
- Without Select locked cells, only unlocked cells can be selected: a click on a locked cell does nothing, an arrow goes on to the next unlocked cell that way (or stays where there is none), and Tab and Enter move to the next unlocked cell along the rows or down the columns, wrapping around, among the cells the sheet has loaded.

The display filter and its sort change no data, so a protected sheet does not limit them, the AutoFilter's buttons included. Data > Filter (`toggleAutoFilter()`, Ctrl+Shift+L) is disabled on a protected sheet whatever its boxes allow, and `./local` refuses `setAutoFilter` there: in Excel the Sort and Use AutoFilter boxes let users use an AutoFilter the sheet has, not set or remove one ([autofilter.md](autofilter.md)).

Not as in Excel: Show Formulas shows a hidden formula, Find's Look in Formulas finds text in it, and Trace Precedents explains it, since the sheet hides it only in the formula bar and the state; the sheet checks only the loaded cells before a change, so a change beyond them reaches the source, whose refusal it shows. Excel's alert for a locked cell is the same on Excel for Windows, but its Protect Sheet dialog words the boxes differently (Windows: "Protect worksheet and contents of locked cells"); the texts here are Excel for Mac's.
