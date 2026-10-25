# Freeze Panes, Reapply, Trace Precedents and Trace Dependents

## Reapply

When a display filter is on, the cell context menu has Excel's Reapply command, also available as Ctrl+Alt+L (on a Mac as well, since the browser reserves Cmd+Option+L) and as a button on the filter's status line at the bottom. Reapply applies the filter conditions and the sort again to the current data through `visibleRows`: rows that no longer match after an edit are hidden, and rows that now match appear. Until then such rows stay as they are, as in Excel. The scroll position is kept, and so is the active cell, unless its row was hidden; the active cell then moves to the row that takes its place.

## Freeze Panes

If the source has `freeze(sheet, panes)` and the sheet is editable, the context menu has Excel's Freeze Panes, Freeze Top Row and Freeze First Column; on a sheet whose workbook has frozen panes, Unfreeze Panes replaces Freeze Panes. Freeze Panes keeps the rows above and the columns to the left of the active cell in view, so with entire columns selected it freezes the columns to the left of the selection, and with entire rows selected the rows above it. With A1 active it freezes the rows and columns above and to the left of the middle of the viewport, where Excel splits the window. Each command is a single undo step in the source, and the sheet reloads afterwards unless the source has sent `reload`.

The lock button on the formula bar only controls how the workbook's frozen panes are displayed: it hides or shows them in this browser and does not change the workbook. By default the sheet does not display frozen panes that take more than half of the viewport. Panes frozen from the menu are displayed regardless of the lock's earlier state and of how much of the viewport they take, and Unfreeze Panes clears the lock's saved choice. Without `freeze` these menu items are absent and the lock works as before. They are also absent under a display filter, because the rows on screen are then not the sheet's rows in order.

The same commands can be run from the host's buttons through the sheet's methods ([api.md](api.md#sheetviewapi)). The example uses `memorySource`, which has `freeze` and `explain`: select C4 and click Trace precedents to see its precedents.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { SheetView, memorySource, type SheetViewApi } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({
  sheets: [
    {
      name: 'Budget',
      cells: {
        A1: 'Item', B1: 'Price', C1: 'Sum',
        A2: 'Desks', B2: 250, C2: { value: 1000, formula: 'B2*Rates!B1' },
        A3: 'Chairs', B3: 90, C3: { value: 360, formula: 'B3*Rates!B1' },
        A4: 'Total', C4: { value: 1360, formula: 'SUM(C2:C3)' },
      },
    },
    { name: 'Rates', cells: { A1: 'Quantity', B1: 4 } },
  ],
})
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <button @click="sheet?.freezePanes('topRow')">Freeze the top row</button>
  <button @click="sheet?.freezePanes('unfreeze')">Unfreeze</button>
  <button @click="sheet?.explain()">Trace precedents</button>
  <SheetView ref="sheet" :source="source" style="height: 360px" />
</template>
```

## Trace Precedents

If the source has `explain(sheet, cell)`, the context menu has Trace Precedents. Instead of drawing arrows as Excel does, the sheet opens a panel to the right of the grid (below it on a phone) with the source's tree of the active cell's precedents. While open, the panel follows the active cell: it requests the tree again 150 ms after the active cell stops moving or its value changes, and when the source reports changes. Since `explain` takes no abort signal, a response that arrives after a newer request is discarded.

Each line is a node, and each two-space indent is one level deeper, as in the text of xlcalc's `explain` (`'Q 4'!A2 = 60  ← Data!A4*2`). References in it to cells, columns or rows on a single sheet are links; a reference without a sheet name refers to the sheet of its line's node. Clicking a link selects its range, like Ctrl+click on a reference in the formula bar, and the panel then shows the precedents of that cell. The close button (✕) and Esc close the panel. `./local` returns the session's `explain` in the workbook's locale, and `memorySource` builds a tree from the formulas as given, with xlcalc's English wording. The panel's width, and its height on a phone, are `--xs-explain-w` (340 px) and `--xs-explain-h` (45 %).

## Trace Dependents

If the source has `dependents(sheet, range)`, the sheet's `traceDependents()` method draws Excel's tracer arrows from the active cell to the cells whose formulas refer to it directly: a dot in the traced cell, a blue line and an arrowhead at each dependent on the same sheet, a red one from a cell whose value is an error, and a dashed black line to a worksheet icon next to the cell when formulas on other sheets refer to it (one icon for all of them, as in Excel). Each further call adds the next level: the dependents of the cells the arrows already reach whose own dependents are not drawn yet. When there is no further level, nothing is added and the method returns `false`; on a cell no formula refers to it also shows Excel's message. A call traces at most 200 cells; what the source leaves out of its response (`more`) is not drawn.

The arrows are drawn in a layer over the cells that scrolls with the grid. An arrow between cells of the scrolling part goes under the frozen panes and the headers, and one with an end in the frozen panes goes over them and follows them, clipped to the cells' area. Each sheet keeps its own arrows. `removeArrows()` is Excel's Remove Arrows and removes every arrow; as in Excel, any change of cells from the sheet (an edit, a paste, a cleared range, a structure operation such as inserting rows) removes them too, and formatting does not. Like Excel, the sheet traces from the active cell, not from each cell of a selected range. The colors are `--xs-trace`, `--xs-trace-error` and `--xs-trace-sheet`.

Excel's Trace Dependents is a ribbon command without a key, and its cell context menu has no such item, so the sheet offers it through the method for a host's own button. `./local` answers `dependents` from the session's dependency graph; `memorySource` has no engine and no `dependents`.

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { SheetView, type SheetViewApi } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({
  sheets: [
    { name: 'Rates', cells: { A1: { value: 'Rate' }, B1: { value: 0.2 }, A2: { value: 'Base' }, B2: { value: 100 }, A3: { value: 'Tax' }, B3: { formula: 'B1*B2' }, A4: { value: 'Total' }, B4: { formula: 'B2+B3' } } },
    { name: 'Report', cells: { A1: { value: 'Tax' }, B1: { formula: 'Rates!B3' } } },
  ],
})
onUnmounted(() => void source.close())
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <!-- select B1, then click twice: B3, then B4 and the icon for Report -->
  <button @click="sheet?.traceDependents()">Trace dependents</button>
  <button @click="sheet?.removeArrows()">Remove arrows</button>
  <SheetView ref="sheet" :source="source" :locale="source.locale" style="height: 360px" />
</template>
```
