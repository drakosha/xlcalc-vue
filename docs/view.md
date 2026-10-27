# Freeze Panes, zoom, Watch Window, Reapply, Show Formulas, Trace Precedents and Trace Dependents

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

## Show Formulas

Ctrl+` (Control+` on a Mac too, as in Excel for Mac; the key left of 1, with any keyboard layout) is Excel's Show Formulas, and pressing it again shows the values. The sheet's methods do the same for a host's button: `showFormulas()` toggles it, `showFormulas(true)` and `showFormulas(false)` set it, and the `state` event's `showFormulas` tells whether it is on, so the button can stay pressed. While it is on, every cell with a formula shows its formula as the window gives it, in the source's locale, with `=` (`=SUM(B2:B9)`, `=СУММ(B2:B9)`), aligned as text (left unless the cell's own alignment says otherwise) and overflowing into empty neighbors as text does. A constant is shown without its number format: a number as the formula bar writes it in General, in the workbook's locale (a date as its serial number, `45688`; `1234.5` for `1,234.50`), through `entryText` of `xlcalc/entry`, a separate chunk that loads when the view is first turned on; text as stored; a boolean or an error as displayed. The number format's color is not applied. Columns are drawn twice as wide, as Excel draws them in this view, while the workbook's widths stay as they are: a width dragged with the mouse meanwhile is stored as half the drawn one, and with the view off the column is drawn at that stored width.

As in Excel, it is a setting of each sheet: another sheet shows its values until it gets its own, and a sheet displayed again keeps its setting. The source has no place for it, so the sheet keeps it as it keeps the frozen panes choice, in the browser's storage under `storageKey` (otherwise for the component's lifetime). Editing works as before: F2 opens the cell's own text, and an entry shows as its formula once it is in. The selection's formula ranges are highlighted on the grid in this view as in the other.
## Zoom

The sheet is drawn at Excel's zoom, from 10% to 400%, and each sheet keeps its own, as each of Excel's sheet views does. Everything in the grid scales together: cells, fonts, borders, row numbers and column letters, frozen panes, the selection, the editor and the boxes over cells (a validation list, its input message). The Formula AutoComplete list and the argument tooltip keep their size, as in Excel, and so does the fill handle, so it can still be grabbed zoomed out. Gridlines and thin borders stay one screen pixel wide when zoomed out. Zoom changes only the view: column widths and row heights stay as the workbook has them, and nothing is sent to the source.

The zoom changes by:

- Ctrl+wheel over the grid (on a Mac too): each notch zooms in or out by 15%. A trackpad's pinch, which the browser sends as Ctrl+wheel (Safari as its gesture events), zooms in proportion to the gesture about the pointer.
- Two fingers on a touch screen: the sheet zooms about their middle, and the page itself does not.
- The status bar, at the right end of the bottom bar on a desktop: Zoom Out (−) and Zoom In (+) go to the next multiple of 10%; the slider covers 10% to 100% on its left half and 100% to 400% on its right, with 100% at the middle; the percentage opens Excel's Zoom dialog.
- The Zoom dialog: 200%, 100%, 75%, 50%, 25%, Fit selection and Custom (10 to 400). Fit selection is Excel's Zoom to Selection: the zoom at which the selection fills the viewport, with the selection scrolled to its top left corner.

The wheel, the slider and the dialog keep the top left cell in place; a pinch keeps the cell under it. A host sets a sheet's zoom with `setZoom(percent)`, or Zoom to Selection with `setZoom('selection')`, and opens a sheet at its zoom with the `zoom` prop, keyed by sheet name; a sheet the prop does not name opens at the workbook's saved zoom (`SheetInfo.zoom`, which `./local` reads from the xlsx sheet view), otherwise at 100%. The `zoom` event reports what the user changes (a pinch once it ends), so the host can keep it and pass it back; `state.zoom` is the displayed sheet's zoom, absent at 100%.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { SheetView, memorySource, type SheetViewApi, type SheetZoomEvent } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({ sheets: [{ name: 'Budget', cells: { A1: 'Item', B1: 'Price', A2: 'Desks', B2: 250 } }] })
const sheet = ref<SheetViewApi | null>(null)
// The zoom by sheet as the user left it, to open the sheets at it again
const zooms = ref<Record<string, number>>({ Budget: 150 })
const onZoom = (e: SheetZoomEvent) => (zooms.value = { ...zooms.value, [e.sheet]: e.zoom })
</script>

<template>
  <button @click="sheet?.setZoom(100)">100%</button>
  <button @click="sheet?.setZoom('selection')">Zoom to selection</button>
  <SheetView ref="sheet" :source="source" :zoom="zooms" style="height: 360px" @zoom="onZoom" />
</template>
```

The grid's content is drawn under CSS `zoom`, which current Chrome, Safari and Firefox support as standardized: its geometry stays in the workbook's pixels, and the pointer's page coordinates are divided by the zoom where the sheet finds the cell under the pointer itself (a drag of the fill handle, a range pointed at in a formula, a column edge).

## Watch Window

Excel's Watch Window lists cells whose values stay in view whatever sheet is displayed. The host opens it with `watchWindow()` (and closes it with `watchWindow(false)`); it floats over the sheet within the component's box, at its top right corner (over the lower part of the grid on a phone), is moved by its title bar, resized by its corner, and closed by its close button.

Add Watch adds each cell of the selection, by rows, that is not watched yet, since Excel has one watch per cell; at most 1,000 cells are watched. The list has Excel's columns: Book, Sheet, Name, Cell, Value (the displayed value) and Formula (with `=`, in the source's locale). Book and Name stay empty: the sheet does not know the workbook's name, and does not look up a defined name for a watched cell. A click chooses a row, Ctrl+click (Cmd+click on a Mac) adds a row to the chosen ones or takes it out, Shift+click chooses the rows from the last one clicked, and the arrow keys move the choice. Delete Watch, or Delete, removes the chosen rows. A double click on a row, or Enter, goes to its cell, on its sheet.

The values are read through the source's `window`, one window over each sheet's watched cells (a window per cell when that would exceed `maxCells`), only while the window is open. They are read again when cells change (an edit's outcome, a `changed` event), when the sheets are read again (a `reload` event, a structure change) and when the list changes; a workbook being recalculated is asked again a second later. The `watches` event reports the list after each Add Watch or Delete Watch, and the `watches` prop replaces it, so the host keeps the list as view state.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { SheetView, memorySource, type SheetViewApi, type SheetWatch } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({
  sheets: [
    { name: 'Summary', cells: { A1: 'Total', B1: { value: 60, formula: 'SUM(Data!B1:B3)' } } },
    { name: 'Data', cells: { B1: 10, B2: 20, B3: 30 } },
  ],
})
const sheet = ref<SheetViewApi | null>(null)
const watches = ref<SheetWatch[]>([{ sheet: 'Summary', cell: 'B1' }])
</script>

<template>
  <button @click="sheet?.watchWindow()">Watch Window</button>
  <SheetView ref="sheet" :source="source" :watches="watches" style="height: 360px" @watches="watches = $event" />
</template>
```

## Trace Precedents

If the source has `explain(sheet, cell)`, the context menu has Trace Precedents. Instead of drawing arrows as Excel does, the sheet opens a panel to the right of the grid (below it in a sheet narrower than 600 px, by the sheet's own width, which the panel does not change) with the source's tree of the active cell's precedents. While open, the panel follows the active cell: it requests the tree again 150 ms after the active cell stops moving or its value changes, and when the source reports changes. Since `explain` takes no abort signal, a response that arrives after a newer request is discarded.

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
