# Hyperlinks, notes, Clear and the default font

## Hyperlinks

With the source's `link` and an editable sheet, Ctrl+K (Cmd+K on a Mac), the sheet's `insertLink()` method and the cell menu's Link open Excel's Insert Hyperlink dialog for the selection. On a cell that has a link the same commands open it as Edit Hyperlink, filled from the link, with Remove Link. The dialog has:

- Text to display: the active cell's text. On an empty cell it follows the address as it is typed, until it is typed in itself, as Excel's does. Over several cells, or on a cell with a formula, it shows `<<Selection in Document>>` and cannot be changed.
- Link to: Existing File or Web Page, with its Address, or Place in This Document, with Type the cell reference (A1 by default) and the places under Cell Reference, which are the workbook's visible sheets. The link's place is the chosen sheet and the reference, written as Excel writes it (`'Q 2'!B3:C4`); a reference that is not a cell or a range ("Reference isn’t valid.") keeps the dialog open. A place of an edited link that is not a cell reference (a defined name) is listed under Defined Names, where the workbook's names are to be listed once the sheet reads them.
- ScreenTip, a field in the dialog rather than Excel's separate ScreenTip box, and OK and Cancel; Enter is OK and Esc is Cancel.

OK writes the link over the selection through `link(sheet, range, link)`, the cells' contents kept, as a single undo step. When Text to display was changed, the sheet first enters the new text into the active cell through `edit`, as typed, and then writes the link: two undo steps where Excel's dialog makes one. An empty Text to display on OK means the address, as in Excel. Excel also gives a new link's cell its Hyperlink cell style; the sheet does not write a style, since it draws every link underlined in `--xs-link`.

The cell menu has Excel's hyperlink items: Link on a cell without a link; Edit Hyperlink, Open Hyperlink and Remove Hyperlink on a cell with one. Open Hyperlink follows the link as a click on it does: a place in the workbook is selected, any other link goes to the host's `openLink` when it listens (or with `hostLinks`), and otherwise a web link opens in a new tab. Remove Hyperlink removes the links over the selection, keeping the cells' text. Right-clicking a web link in a cell, or a long press on it, opens this menu when the sheet can edit links, as Excel's does; otherwise it leaves the browser's own menu for the link, as before. On a read-only sheet, or without `link`, only Open Hyperlink is offered, and Ctrl+K does nothing.

After `link` succeeds the sheet reads the windows over the range again, whether or not the source sends a `repaint`.

## Notes

With the source's `note` and an editable sheet, the cell menu has Excel's New Note on a cell without a note, and Edit Note and Delete Note on a cell with one; Shift+F2 does New Note or Edit Note for the active cell, as in Excel. The note is edited in a box anchored to its cell, Excel's pale yellow note box to the right of the cell (to its left where the viewport has no room), with the author's name in bold on the first line: for a new note the `noteAuthor` prop, which a host sets to the user's name as Excel takes it from its Options, and for an existing note its own author. Enter starts a new line. Esc or a press outside the box saves the note through `note(sheet, cell, note)`, as Excel keeps a note's text once you leave it; an unchanged note sends nothing, and neither does a new one left empty without an author. Delete Note removes the active cell's note through `note(sheet, cell, null)`. Each is a single undo step, after which the sheet reads the windows over the cell again. Notes are displayed as before, with the red indicator and the cell's tooltip; Excel's Show/Hide Note and Show All Notes are not provided.

## Clear Formats and Clear All

Excel has Clear Formats and Clear All under Home > Clear, not in the cell context menu, so the sheet has them as its method `clear(what)` for a host's own ribbon or buttons ([api.md](api.md#sheetviewapi)). With the source's `clear` and an editable sheet, `clear('formats')` removes the styles, number formats and conditional formats over the selection, and `clear('all')` also its contents, hyperlinks, notes and data validation rules, each as a single call to the source and so a single undo step. The selection goes as Excel writes it (`B2:D5`, `B:D`, `2:5`); under a display filter only its visible rows are cleared, as Excel clears only the visible cells of a filtered range, so the call carries the runs of visible rows (`A1:A2`, `A4`). A locked cell in the selection displays its reason and nothing is sent. Since the outcome carries the cells' values and not their formatting, the sheet reads the windows over the cleared ranges again in place, as on a `repaint`, whether or not the source sends one. Delete clears contents only, through `edit`, as before.

## The workbook's default font

A sheet whose `SheetInfo` has `font` (`{ name, size }`, the workbook's Normal style) draws every cell without a font of its own in that font, as Excel draws it: the grid's table takes the font's name and its size in the grid's pixels (11 pt is 12.5 px, the rest in proportion), and the cells inherit it. A cell's own font name or size takes precedence. The row numbers and column letters keep the interface font, `--xs-font`. Without `font` nothing changes: cells are drawn in `--xs-font` and `--xs-font-size`. The row heights are the source's either way (`rowHeights`, `defaultRowHeight`), and overflow into empty neighbors goes by the column widths, so neither depends on the font.

`@xlcalc/vue/local` gives the workbook's font, and Calibri 11 for a workbook without one, since xlcalc measures column widths and `####` in Calibri 11 whatever the font. A browser without Calibri (most outside Windows and Office) draws the cells in a fallback font, whose widths differ, so a value can overflow where xlcalc says it fits. Calibri itself may not be served from a web page, but Carlito, under the SIL Open Font License, has Calibri's widths; a host serves it under Calibri's name, and an installed Calibri still takes precedence:

```css
@font-face {
  font-family: Calibri;
  src: local('Calibri'), url('/fonts/Carlito-Regular.woff2') format('woff2');
}
```
