# Tables

If the source has `table(op)` and the sheet is editable, Ctrl+T and Ctrl+L, Excel's keys for Create Table, open its dialog, and so does the sheet's `insertTable()` method, which a host's Insert Table or Format as Table button calls. The keys are Ctrl on a Mac too, as in Excel for Mac, since the browser keeps Cmd+T and Cmd+L; some browsers keep Ctrl+T for a new tab as well (Chrome on Windows and Linux), and Ctrl+L or the method works there.

The dialog proposes the selection when it has more than one cell (clipped to the data), otherwise the active cell's current region, as Excel's Range.CurrentRegion finds it: the range grown while a row or column next to it, corners included, has a non-empty cell. The range is shown as Excel shows it, `=$A$1:$C$5`, and can be typed with or without `$` and `=`, with the sheet's own name or without one; a range Excel would not take keeps the dialog open with its message. My table has headers is checked when the first row holds text in every column above at least one more row.

OK creates the table through `table({ op: 'addTable', sheet, table })`, named `Table1`, `Table2` and so on: the first number from the count of the workbook's tables on that no table and no defined name has (the word is `tableNamePrefix`, `Таблица` in Russian). With headers, the columns are named after the header cells' text: a number or a date as displayed, an empty header `Column1` (the first number no other column uses), and a name a column to the left already has gets the next free number (`Qty`, `Qty2`), compared case-insensitively, as Excel names them. Without headers, Excel puts a header row of `Column1`, `Column2`… above the data: the sheet first moves the cells from the range's first row down by one row in the range's columns through `structure` (`moveRange`), then adds the table over the range one row taller, so this takes two undo steps where Excel takes one. Without `structure` the box stays checked. The table is drawn with the sheet's table colors (`--xs-table-head-bg`, `--xs-table-line`); the library's table style and Excel's gallery of styles are not offered yet.

`convertToRange()` is Excel's Convert to Range on the table at the active cell: after Excel's question, Yes sends `convertTable` and the cells keep their contents. `deleteTable()` deletes the table at the active cell with its cells (`deleteTable`). Both resolve to `null` when there is no table at the active cell, the source has no `table`, or the answer is No. `./local` implements `table` with xlcalc's table operations; `memorySource` has none.

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { SheetView, type SheetViewApi } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({
  sheets: [
    {
      name: 'Orders',
      cells: {
        A1: { value: 'Item' }, B1: { value: 'Qty' }, C1: { value: 'Qty' },
        A2: { value: 'Pens' }, B2: { value: 10 }, C2: { value: 12 },
        A3: { value: 'Paper' }, B3: { value: 3 }, C3: { value: 5 },
      },
    },
  ],
})
onUnmounted(() => void source.close())
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <!-- Insert Table proposes =$A$1:$C$3 with headers, and names the columns Item, Qty, Qty2 -->
  <button @click="sheet?.insertTable()">Insert table</button>
  <button @click="sheet?.convertToRange()">Convert to range</button>
  <SheetView ref="sheet" :source="source" :locale="source.locale" style="height: 360px" />
</template>
```
