# Rows, columns and sheets

A right click on a row number or a column letter, or a finger held on it for half a second, opens Excel's menu for the selected lines (a line outside a selection of whole lines is selected first): Insert and Delete as many lines as are selected through `structure(op)`, Row Height or Column Width in Excel's units (points, up to 409.5; characters of the default font, as Excel for Mac's box shows them) with Excel's limits, Hide and Unhide (always there, as in Excel), through `layout(change)`; a size of 0 hides the lines. The same on a tab makes that sheet the active one and offers Insert (before it, under Excel's next name: `Sheet4`, `Лист4` with `ruMessages`), Delete (asked first when the sheet has data), Rename (in place on the tab, a double click too), Move Left and Move Right among the shown tabs, Hide, and Unhide with the list of hidden sheets. A refusal shows the source's message (`./local` gives the library's, in the book's locale), and after a change the sheet reads its sheets and windows again unless the source sent `reload`. The sheet shown after is Excel's: a sheet inserted or unhidden, and for one deleted or hidden the next shown sheet, or the one before it at the end. Without `structure` or `layout` their items are gone, without both the browser's own menu stays, and a read-only sheet has none. The boxes of a size, of the question before a delete, of the hidden sheets and of Group and Ungroup are modal: Esc or a press outside closes them, and Tab and Shift+Tab stay inside.

A book with `localSource`, which has `structure`, `layout` and `outline`: right-click a row number, a column letter or a tab for their menus (Unhide on a tab offers the hidden sheet), and press the outline's buttons left of the rows, where South is collapsed.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from 'xlcalc-sheet'
import { localSource } from 'xlcalc-sheet/local'
import 'xlcalc-sheet/style.css'

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
      // Two groups of level 2 inside one of level 1; South's rows are hidden and its summary row keeps the mark
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

A sheet whose `SheetInfo.outline` has groups gets Excel's outline bars: the rows' left of the row numbers, the columns' above the letters, a slot for each level and one more. Each expanded group has a thin bracket along its lines with a tick at the end away from its summary line, and a − button on the summary line (below the group, or above it without `summaryBelow`; right of columns, or left without `summaryRight`); a collapsed group has only its +. The level buttons 1 to N+1 sit in the corner: the rows' across their bar at the letters' height, the columns' down their bar over the numbers. A group's button goes to `outline({ op: 'collapse' | 'expand', sheet, rows: '3:6', level })` (`columns: 'B:D'` for columns), a level button to `outline({ op: 'showLevel', sheet, axis, level })`, and the sheet reads itself again after them unless the source sent `reload`. Without `outline`, or on a read-only sheet, the bars are drawn and their buttons do not press. Under a display filter the rows' bar is left out, as its rows are not the sheet's in order. On a finger the bars' slots are wider and a tap on a bar presses the nearest button within 22 px, a target of 44 px around a button drawn at Excel's size. The menu on numbers and letters has Group and Ungroup over the selected lines, as do Excel's keys, Shift+Alt+Right and Left (Cmd+Shift+K and J on a Mac); over a selection of cells a small box asks rows or columns, rows first, as Excel's does. What Excel does not do is said and not sent: a Group past the seventh level, an Ungroup of lines in no group. The bars' colours and the buttons' size are `--xs-outline-*` variables.

The tabs show with one sheet too and end with "+", which inserts a sheet at the end and shows it (with `structure`); hidden sheets are not on them. A double click on a column's edge fits the column to what it shows through `layout({ columns, fit: true })`, every selected whole column when it is one of them, as Excel does; without `layout` it brings back the book's width. With `layout` a row number's bottom edge drags the row's height with a mouse (a touch screen has Row Height in the line menu: a finger on a number would snap to the edge), drawn as it moves, and the release sets it in points through `layout({ rows, height })`, for every selected whole row when the dragged one is among them, as Excel does; a double click there fits the row (and the selected whole rows it is one of) to what it shows through `layout({ rows, fit: true })`, which `./local` gives to xlcalc's AutoFit of rows. With `./local` a text typed with a line break (Alt+Enter) wraps its cell and fits its row, unless the row's height was set by hand, as Excel does (xlcalc 0.33). Two properties of `SheetView` give back an older grid's ways: `hide-lone-tab` (no tabs for a book of one sheet) and `edge-restores-width` (the double click brings back the book's width).
