# Rows, columns and sheets

## Context menus

Right-clicking a row number or a column letter, or a long press of half a second on it on a touch screen, opens Excel's context menu for the selected rows or columns (a row or column outside a selection of entire rows or columns is selected first):

- Insert and Delete, as many rows or columns as are selected, through `structure(op)`;
- Row Height or Column Width in Excel's units (points, up to 409.5; characters of the default font, as Excel for Mac's dialog shows them) with Excel's limits, and Hide and Unhide (always present, as in Excel), through `layout(change)`; a size of 0 hides the rows or columns.

The same on a sheet tab makes that sheet active and offers Insert (before it, with Excel's next default name: `Sheet4`, or `Лист4` with `ruMessages`), Delete (with a confirmation when the sheet has data), Rename (in place on the tab, also by double-clicking it), Move Left and Move Right among the visible tabs, Hide, and Unhide with the list of hidden sheets. If the source rejects a command, its message is displayed (`./local` returns the library's message in the workbook's locale), and after a change the sheet reads its sheets and windows again unless the source has sent `reload`. The sheet displayed afterwards is the one Excel displays: an inserted or unhidden sheet, and after a sheet is deleted or hidden, the next visible sheet, or the previous one if it was the last. Without `structure` or `layout` the corresponding items are absent; without both, the browser's own context menu remains, and a read-only sheet has none. The dialogs for sizes, the delete confirmation, the hidden sheets, and Group and Ungroup are modal: Esc or a click outside closes them, and Tab and Shift+Tab stay inside.

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
  <SheetView :source="source" :locale="source.locale" style="height: 360px" />
</template>
```

## Outline

A sheet whose `SheetInfo.outline` has groups gets Excel's outline bars: the row bar to the left of the row numbers and the column bar above the column letters, with a slot for each level plus one. Each expanded group has a thin bracket along its rows or columns with a tick at the end away from its summary row or column, and a − button on the summary row or column (below the group, or above it without `summaryBelow`; to the right of the columns, or to the left without `summaryRight`); a collapsed group has only its + button. The level buttons 1 to N+1 are in the corner: the rows' buttons run across their bar at the height of the column letters, the columns' buttons run down their bar above the row numbers. A group's button calls `outline({ op: 'collapse' | 'expand', sheet, rows: '3:6', level })` (`columns: 'B:D'` for columns), a level button calls `outline({ op: 'showLevel', sheet, axis, level })`, and the sheet reloads afterwards unless the source has sent `reload`. Without `outline`, or on a read-only sheet, the bars are drawn but their buttons do nothing. Under a display filter the row bar is not shown, because the displayed rows are not the sheet's rows in order. On a touch screen the bars' slots are wider, and a tap on a bar presses the nearest button within 22 px, a 44 px target around a button drawn at Excel's size. The context menu of row numbers and column letters has Group and Ungroup for the selected rows or columns, as do Excel's shortcuts Shift+Alt+Right and Shift+Alt+Left (Cmd+Shift+K and Cmd+Shift+J on a Mac); with a range of cells selected, a small dialog asks whether to group rows or columns, rows first, as Excel's does. Commands that Excel does not perform are reported and not sent: a Group beyond the seventh level, an Ungroup of rows or columns that are in no group. The bars' colors and the buttons' size are set by the `--xs-outline-*` variables.

## Tabs and sizes

The sheet tabs are shown even for a single sheet and end with a "+" button, which inserts a sheet at the end and displays it (with `structure`); hidden sheets have no tabs. Double-clicking a column header's edge fits the column to its contents through `layout({ columns, fit: true })`, or every selected entire column when the column is one of them, as Excel does; without `layout` it restores the workbook's width. With `layout`, the bottom edge of a row number drags the row's height with a mouse (on a touch screen, Row Height is in the row menu, since a finger on a row number would snap to the edge); the row is redrawn as it moves, and the release sets the height in points through `layout({ rows, height })`, for every selected entire row when the dragged row is among them, as Excel does. Double-clicking that edge fits the row (and the selected entire rows it belongs to) to its contents through `layout({ rows, fit: true })`, which `./local` passes to xlcalc's row AutoFit. With `./local`, text typed with a line break (Alt+Enter) wraps its cell and fits its row, unless the row height was set manually, as in Excel (xlcalc 0.33). Two `SheetView` props reproduce an older grid's behavior: `hide-lone-tab` (no tabs for a workbook with one sheet) and `edge-restores-width` (double-clicking a column edge restores the workbook's width).
