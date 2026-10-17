# Clipboard and fill handle

Control+C (Cmd+C on a Mac) puts the selection on the clipboard as plain text, as Excel puts it (xlcalc's `formatClipboard`): fields at tabs, rows at CRLF with one at the end (none on a Mac), a field with a tab or a line break in quotes, hidden columns and rows left out. With the source's `copyText` (`./local` has it) the copy event puts the shown text at once and then the source's own text through the browser's clipboard API: values by their formats without the column's clipping, as Excel copies them (a narrow `##` copies as its date); a browser that refuses the clipboard API keeps the shown text. A paste of either text is this copy. Without `copyText`, or under the display filter, a cell copies the text it shows, and marks the range with Excel's moving dashed border; Control+X does the same as a cut. With `structure` on an editable sheet, Control+V of that same text, or a paste whose clipboard the browser does not give, pastes the range itself: a copy through `structure({ op: 'copyRange', from, to })`, as tiles into a selection whose size the copy divides and else once from the selection's top-left cell, and a cut through `moveRange` to the selection's top-left cell, on another sheet of the book too. The source does the operation, so formulas move as Excel moves them (`./local` hands it to xlcalc's `editStructure`). The pasted range is selected, and a refusal shows the source's message. After a paste a copy's border stays until Esc, a cell's field opened, a change of cells (an edit, a fill, Delete, an undo) or another copy; a cut's goes once it is pasted. Once the border is gone, its text is not pasted again, as Excel has nothing to paste then. Any other text is pasted as before (`paste`, else `edit`), and formatting and lines leave the border on. Without `structure` a copy is pasted as its text. Under the display filter a copy is the text of the shown rows, unmarked, and there is no cut.

A sheet to try them on, with `localSource`, which has `structure` and `copyText`: select A1:D2 and drag the fill handle down to row 6 for the series and the formula moved, or copy E2 and paste it over E3:E6.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from 'xlcalc-sheet'
import { localSource } from 'xlcalc-sheet/local'
import 'xlcalc-sheet/style.css'

const source = localSource({
  sheets: [
    {
      name: 'Series',
      cells: {
        A1: { value: 1 }, B1: { value: 'Mon' }, C1: { value: 'Q1' }, D1: { value: 'Item1' }, E1: { formula: 'A1*10' },
        A2: { value: 2 }, B2: { value: 'Tue' }, C2: { value: 'Q2' }, D2: { value: 'Item2' }, E2: { formula: 'A2*10' },
      },
    },
  ],
})
onUnmounted(() => void source.close())
</script>

<template>
  <SheetView :source="source" :locale="source.locale" style="height: 320px" />
</template>
```

A right click on a cell, or a finger held on it, opens Excel's menu for cells over the selection (a cell outside it is selected first): Cut, Copy and Paste, 44 px high on a finger, then the commands of [view.md](view.md) the source allows. This menu and the menus of the row numbers, column letters and tabs take the keyboard as Excel's: the first enabled item has the focus, Down and Up go round the enabled ones, Home and End to the first and the last, Enter or Space chooses, Esc and Tab close the menu and give the focus back. They go through the browser's clipboard API. Paste reads the clipboard, which the browser may ask the person to allow; where it will not, Paste pastes only the sheet's own copy and otherwise says to use Control+V. A read-only sheet's menu has Copy and Trace Precedents alone.

Excel's fill handle is the small square at the selection's bottom-right corner, on an editable sheet with `structure`, but not under the display filter, over whole columns or rows, or while a field is open. Dragged down, up, right or left, its frame grows cell by cell along the axis the pointer went furthest, and scrolls the grid at its edge; on release the frame is filled from the selection through `structure({ op: 'fillRange', sheet, range, to })`, Excel's AutoFill, one step of undo, and the frame is selected. The source does the fill (`./local` hands it to xlcalc's `editStructure`): series go on (1, 2 → 3, 4; dates by day; Item1 → Item2; Q1, Mon, Jan), the rest is copied with formulas moved, a part of the selection fills the frame's end, and a drag up or left goes backwards. Released with Ctrl (Option on a Mac) it sends `type: 'series'`, which makes a single number go on by 1 instead of being copied. Esc while it is dragged drops the drag, as in Excel, and the release fills nothing; a drag back inside the selection does nothing (Excel clears the cells it leaves). A double click on it fills down as far as the data of the column left of the selection goes, else of the column right of it. On a finger its target is 44 px around a square drawn small, and a tap on it picks the cell under it. A refusal (a merged area, a source whose library has no `fillRange`) shows the source's message. The border's and the frame's colours are `--xs-copy-line` and `--xs-fill-frame`.
