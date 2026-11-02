# Rows, columns and sheets

## Context menus

Right-clicking a row number or a column letter, or a long press of half a second on it on a touch screen, opens Excel's context menu for the selected rows or columns (a row or column outside a selection of entire rows or columns is selected first):

- Insert and Delete, as many rows or columns as are selected, through `structure(op)`;
- Row Height or Column Width in Excel's units (points, up to 409.5; characters of the default font, as Excel for Mac's dialog shows them) with Excel's limits, and Hide and Unhide (always present, as in Excel), through `layout(change)`; a size of 0 hides the rows or columns.

The same on a sheet tab makes that sheet active and offers Insert (before it, with Excel's next default name: `Sheet4`, or `Лист4` with `ruMessages`), Delete (with a confirmation when the sheet has data), Rename (in place on the tab, also by double-clicking it), Move Left and Move Right among the visible tabs, Tab Color with the source's `tabColor` ([below](#tab-color)), Hide, Unhide with the list of hidden sheets, and with the source's `protect` Protect Sheet… or, on a protected sheet, Unprotect Sheet ([protection.md](protection.md)). On a protected sheet the row and column menus keep their items and disable those its protection does not allow, as Excel does. If the source rejects a command, its message is displayed (`./local` returns the library's message in the workbook's locale), and after a change the sheet reads its sheets and windows again unless the source has sent `reload`. The sheet displayed afterwards is the one Excel displays: an inserted or unhidden sheet, and after a sheet is deleted or hidden, the next visible sheet, or the previous one if it was the last. Without `structure` or `layout` the corresponding items are absent (the tab menu then has only Tab Color and Protect Sheet, with their methods); without any of them, the browser's own context menu remains, and a read-only sheet has none. The dialogs for sizes, the delete confirmation, the hidden sheets, and Group and Ungroup are modal: Esc or a click outside closes them, and Tab and Shift+Tab stay inside.

A workbook over `localSource`, which has `structure`, `layout` and `outline`: right-click a row number, a column letter or a sheet tab to open its menu (Unhide on a tab offers the hidden sheet), and click the outline buttons to the left of the rows, where South is collapsed.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({
  sheets: [
    {
      name: 'Sales',
      cells: {
        A1: { value: 'Region' }, B1: { value: 'Sales' },
        A2: { value: 'North A' }, B2: { value: 10 },
        A3: { value: 'North B' }, B3: { value: 20 },
        A4: { value: 'North' }, B4: { formula: 'SUBTOTAL(9,B2:B3)' },
        A5: { value: 'South A' }, B5: { value: 5 },
        A6: { value: 'South B' }, B6: { value: 15 },
        A7: { value: 'South' }, B7: { formula: 'SUBTOTAL(9,B5:B6)' },
        A8: { value: 'Total' }, B8: { formula: 'SUBTOTAL(9,B2:B7)' },
      },
      // Two level-2 groups inside a level-1 group; South's rows are hidden and its summary row keeps the collapsed flag
      rows: {
        '2:3': { outlineLevel: 2 },
        '4': { outlineLevel: 1 },
        '5:6': { outlineLevel: 2, hidden: true },
        '7': { outlineLevel: 1, collapsed: true },
      },
      columns: { A: { width: 14 } },
    },
    { name: 'Archive', hidden: true, cells: { A1: { value: 'Last year' } } },
  ],
})
onUnmounted(() => void source.close())
</script>

<template>
  <SheetView :source="source" style="height: 360px" />
</template>
```

## Outline

A sheet whose `SheetInfo.outline` has groups gets Excel's outline bars: the row bar to the left of the row numbers and the column bar above the column letters, with a slot for each level plus one. Each expanded group has a thin bracket along its rows or columns with a tick at the end away from its summary row or column, and a − button on the summary row or column (below the group, or above it without `summaryBelow`; to the right of the columns, or to the left without `summaryRight`); a collapsed group has only its + button. The level buttons 1 to N+1 are in the corner: the rows' buttons run across their bar at the height of the column letters, the columns' buttons run down their bar above the row numbers; a sheet without headings has no level buttons, as in Excel for Mac. A group's button calls `outline({ op: 'collapse' | 'expand', sheet, rows: '3:6', level })` (`columns: 'B:D'` for columns), a level button calls `outline({ op: 'showLevel', sheet, axis, level })`, and the sheet reloads afterwards unless the source has sent `reload`. Without `outline`, or on a read-only sheet, the bars are drawn but their buttons do nothing. Under a display filter the row bar is not shown, because the displayed rows are not the sheet's rows in order. On a touch screen the bars' slots are wider, and a tap on a bar presses the nearest button within 22 px, a 44 px target around a button drawn at Excel's size. The context menu of row numbers and column letters has Group and Ungroup for the selected rows or columns, as do Excel's shortcuts Shift+Alt+Right and Shift+Alt+Left (Cmd+Shift+K and Cmd+Shift+J on a Mac); with a range of cells selected, a small dialog asks whether to group rows or columns, rows first, as Excel's does. Commands that Excel does not perform are reported and not sent: a Group beyond the seventh level, an Ungroup of rows or columns that are in no group. The bars' colors and the buttons' size are set by the `--xs-outline-*` variables.

## Tab Color

A sheet whose `SheetInfo` has `tabColor` (`#RRGGBB`, xlcalc's `SheetInput.tabColor`, xlsx `<tabColor>`) gets Excel's colored tab: another tab is filled with the color, its name in black on a light color and in white on a dark one (by the color's perceived brightness; `--xs-tab-text-dark` and `--xs-tab-text-light`), and the active tab keeps its own background with a thick line of the color under the name, as Excel underlines the selected colored tab.

With the source's `tabColor` on an editable sheet, the tab menu has Tab Color after Move Right, as Excel lists it before Hide. It opens a small box with No Color and Excel's palette, the tab's color shown pressed; a click on a color or No Color applies it at once, as Excel's submenu does, through `tabColor(sheet, color)` (`null` for No Color), and the sheet reads its sheets again unless the source has sent `reload`. `./local` sets the color through xlcalc's `setLayout` (`{ sheet, tabColor }`), a single undo step that changes no value, and writes it to the file; `memorySource` passes a sheet's `tabColor` to its tab and has no `tabColor` method.

## Workbook Statistics

The sheet's `workbookStatistics()` is Excel's Review > Workbook Statistics for a ribbon: it asks the source's `workbookStatistics(sheet)` for the displayed sheet and opens Excel's window with the answer, xlcalc's `WorkbookStatistics` (`./local`: the session's `getWorkbookStatistics`). The window has Current Sheet: and Workbook:, each a list of rows of a label and a value, and OK; Esc closes it too. It shows the rows Excel for Mac shows, as xlcalc measured them: the sheet's End of sheet where it has one (the last row and column of data and notes, not the used range), Cells with data, Tables and Formulas always, any other count (PivotTables, Charts, Images, Form Controls, Objects, Comments, Notes) only above 0, and the workbook's Sheets, but not its Notes, which Excel for Mac never shows. A count the workbook does not hold (`null`: xlcalc has no PivotTables, charts, pictures, form controls, objects or threads of comments) is not shown. A number has no grouping, as in Excel (`5000`). The method returns `false` without the source's `workbookStatistics` or while no sheet is displayed; a failed read goes to the message line. Its texts are `statisticsTitle`, `statisticsSheet`, `statisticsWorkbook`, `statisticsEnd`, `statisticsSheets`, `statisticsCells`, `statisticsTables`, `statisticsPivotTables`, `statisticsFormulas`, `statisticsCharts`, `statisticsImages`, `statisticsFormControls`, `statisticsObjects`, `statisticsComments` and `statisticsNotes`, with `ok`; the title, the sections and the rows Excel for Windows showed are its labels in English and in Russian (`Статистика книг`, `в книге:`), and the rows it did not show on the measured workbook keep Excel's terms (`Сводные таблицы`, `Заметки`).

## Tabs and sizes

The sheet tabs are shown even for a single sheet and end with a "+" button, which inserts a sheet at the end and displays it (with `structure`); hidden sheets have no tabs. Double-clicking a column header's edge fits the column to its contents through `layout({ columns, fit: true })`, or every column of the selection when the column is in it (entire columns or a range of cells), as Excel does; without `layout` it restores the workbook's width. With `layout`, the bottom edge of a row number drags the row's height with a mouse (on a touch screen, Row Height is in the row menu, since a finger on a row number would snap to the edge); the row is redrawn as it moves, and the release sets the height in points through `layout({ rows, height })`, for every selected entire row when the dragged row is among them, as Excel does. Double-clicking that edge fits the row (and every row of the selection it is in, entire rows or a range of cells, as Excel's AutoFit) to its contents through `layout({ rows, fit: true })`, which `./local` passes to xlcalc's row AutoFit. With `./local`, text typed with a line break (Alt+Enter) wraps its cell and fits its row, unless the row height was set manually, as in Excel (xlcalc 0.33). Two `SheetView` props reproduce an older grid's behavior: `hide-lone-tab` (no tabs for a workbook with one sheet) and `edge-restores-width` (double-clicking a column edge restores the workbook's width).
