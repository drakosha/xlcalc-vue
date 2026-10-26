# Clipboard and fill handle

## Copy, cut and paste

Ctrl+C (Cmd+C on a Mac) copies the selection to the clipboard as plain text in Excel's format (xlcalc's `formatClipboard`): fields separated by tabs, each row ending with CRLF, the last one included (no trailing line break on a Mac), fields that contain a tab or a line break in quotes, and hidden rows and columns omitted. If the source has `copyText` (`./local` does), the copy event puts the displayed text on the clipboard immediately and then replaces it with the source's text through the browser's Clipboard API: values formatted by their number formats and not clipped to the column width, as Excel copies them (a date displayed as `##` in a narrow column is copied as the date). If the browser denies the Clipboard API, the displayed text stays on the clipboard. Pasting either text counts as pasting this copy. Without `copyText`, or under a display filter, each cell is copied as its displayed text.

The copied range is marked with Excel's moving dashed border; Ctrl+X does the same for a cut. If the source has `structure` and the sheet is editable, Ctrl+V with that same text on the clipboard, or a paste whose clipboard contents the browser does not provide, pastes the range itself, to another sheet of the workbook as well:

- a copy goes through `structure({ op: 'copyRange', from, to })`, tiled across the selection when the copied range divides the selection's size, otherwise pasted once at the selection's top-left cell;
- a cut goes through `moveRange` to the selection's top-left cell.

The source performs the operation, so formulas are adjusted as Excel adjusts them (`./local` passes it to xlcalc's `editStructure`). The pasted range is selected, and a rejection displays the source's message. After a paste, a copy's border remains until Esc, opening a cell editor, a change to cells (an edit, a fill, Delete, an undo) or another copy; a cut's border disappears once the cut is pasted. Once the border is gone, pasting its text does nothing, since Excel has nothing to paste at that point. Any other text is pasted as before (through `paste`, otherwise `edit`), and changes to formatting or to rows and columns leave the border in place. Without `structure` a copy is pasted as text. Under a display filter a copy contains the text of the visible rows, the range is not marked, and there is no cut.

Try them on a sheet over `localSource`, which has `structure` and `copyText`: select A1:D2 and drag the fill handle down to row 6 to continue the series and fill the formula down, or copy E2 and paste it into E3:E6.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

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

## Context menu

Right-clicking a cell, or a long press on it on a touch screen, opens Excel's cell context menu for the selection (a cell outside the selection is selected first): Cut, Copy and Paste (44 px high on a touch screen), followed by the commands from [view.md](view.md) that the source supports, then the note and hyperlink items of [links-notes.md](links-notes.md). This menu and the context menus of row numbers, column letters and sheet tabs handle the keyboard as Excel's do: the first enabled item has focus, Down and Up cycle through the enabled items, Home and End go to the first and the last, Enter or Space chooses an item, and Esc or Tab closes the menu and returns focus. The menu's Cut, Copy and Paste use the browser's Clipboard API. Paste reads the clipboard, which the browser may ask the user to allow; where it does not allow it, Paste pastes only the sheet's own copy and otherwise tells the user to press Ctrl+V. On a read-only sheet the menu has only Copy, Trace Precedents and, on a cell with a link, Open Hyperlink.

## Fill handle

Excel's fill handle is the small square at the bottom-right corner of the selection. It appears on an editable sheet whose source has `structure`, but not under a display filter, when entire columns or rows are selected, or while a cell editor is open. Dragging it down, up, right or left extends its outline cell by cell along the axis in which the pointer has moved farthest, scrolling the grid at its edge. On release, the outlined range is filled from the selection through `structure({ op: 'fillRange', sheet, range, to })`, Excel's AutoFill, as a single undo step, and the outlined range is selected. The source performs the fill (`./local` passes it to xlcalc's `editStructure`): series continue (1, 2 → 3, 4; dates by day; Item1 → Item2; Q1, Mon, Jan), other values are copied with formulas adjusted, a partial copy of the selection fills the end of the range, and dragging up or left fills backwards. Releasing with Ctrl held (Option on a Mac) sends `type: 'series'`, which makes a single number increase by 1 instead of being copied. Esc during the drag cancels it, as in Excel, and the release then fills nothing; dragging back inside the selection does nothing (Excel clears the cells left out). Double-clicking the fill handle fills down as far as the data in the column to the left of the selection extends, otherwise in the column to the right. On a touch screen its target is 44 px around the small square, and a tap on it selects the cell underneath. If the source rejects the fill (a merged range, a source whose library has no `fillRange`), its message is displayed. The colors of the copy border and the fill outline are `--xs-copy-line` and `--xs-fill-frame`.
