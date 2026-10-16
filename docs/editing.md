# Editing, keys and the source

Editing takes Excel's keys: typing over a cell, F2 or a double click opens the field, Enter, Tab and their Shift commit and move (inside a selection of several cells, which stays, as Excel's: down a column and on to the next one's top, Tab along a row and on to the next row's start, round past the end), Esc cancels, Alt+Enter breaks the line, Control+Shift+U (or the arrow at the formula bar's right end) expands the formula bar to five lines and collapses it to one, as Excel's (collapsed, a cell of several lines shows its first line and the grid stays put), Delete and Backspace clear the selection, Control+Page Down and Page Up show the next and the previous shown sheet, Control and an arrow go where the source's `end(sheet, cell, direction)` says, as Excel's Range.End (with Shift the selection grows there; without `end`, or under the display filter, to the edge of the data the sheet knows). Without `edit` the sheet is read-only.

A book computed in the browser, to try the keys on: type into B2 or C3 and D4 follows, F2 edits a cell in place, Control+Return over B2:C3 enters one text into all four, and `=(1+2` brings Excel's typo window.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from 'xlcalc-sheet'
import { localSource } from 'xlcalc-sheet/local'
import 'xlcalc-sheet/style.css'

const source = localSource({
  sheets: [
    {
      name: 'Prices',
      cells: {
        A1: { value: 'Item' }, B1: { value: 'Price' }, C1: { value: 'Qty' }, D1: { value: 'Sum' },
        A2: { value: 'Pens' }, B2: { value: 1.5 }, C2: { value: 10 }, D2: { formula: 'B2*C2' },
        A3: { value: 'Paper' }, B3: { value: 4.25 }, C3: { value: 3 }, D3: { formula: 'B3*C3' },
        A4: { value: 'Total' }, D4: { formula: 'SUM(D2:D3)' },
      },
    },
  ],
})
onUnmounted(() => void source.close())
</script>

<template>
  <SheetView :source="source" :locale="source.locale" style="height: 360px" />
</template>
```

## The formula bar

The formula bar is a field that always holds the active cell's text, the one the cell's field opens with (a formula with its `=`), as Excel's; while a field is open in the cell it shows the draft. A click in it puts the caret where it was pressed and starts the edit there, as F2 does, a drag starts it with the text dragged over selected, and a click on it while the cell's field is open moves the draft there with the caret where it was pressed. The field's focus starts the edit too, so what is typed into it always goes to the cell. A formula's references are coloured as their frames on the grid, the draft's while one is typed. Outside an edit Ctrl+click on a reference (Cmd+click on a Mac) selects its range, on another sheet too, as a code editor goes to a definition, and the field's tooltip says so (`SheetMessages`' `goRefHint`). While Ctrl or Cmd is held, the reference under the mouse is underlined in its colour and the cursor over it is a hand, whichever came first, the key or the mouse; both go when the key is let go, the mouse leaves the reference or the window loses the focus. Under a finger a long press on a reference (half a second, as the long press that opens the menus; a finger that moves cancels it) goes to its range instead, starts no edit and holds back the browser's own menu and selection, while a tap puts the caret and edits as before; the field has no tooltip there. On a read-only sheet (without `edit`), on a locked cell or with nothing selected the field is read-only: its text can be selected and copied, keys start no edit, and Ctrl+click or a long press on a reference still goes to it. Left of the field are Excel's Cancel (✕, `SheetMessages`' `cancel`) and Enter (✓, `barEnter`), on only while a field is open in the cell or the bar: Cancel is Esc, Enter enters the draft and the active cell stays, as Control+Return on one cell; a formula that does not read brings the same question as Enter. A press on either keeps the field's focus until the button acts. After them is fx, Excel's Insert Function ([autocomplete.md](autocomplete.md#insert-function)). The colours are a layer over the field's transparent text, painted with the CSS Custom Highlight API, so the layer's text lies on the field's to the pixel; a browser without the API gets coloured spans, which may lie a fraction of a pixel off.

## Point mode

While a formula is typed, in the cell's field or the formula bar, the mouse points at references as in Excel's Point mode where xlcalc's `refAt`, measured in Excel, says a reference goes in: after the formula's first character (`=`, or a leading `+` or `-`), an operator, `(`, the book's list separator (`;` in ru-RU), `:`, a sheet's `!`, a point, or a space after an operand (Excel's intersection). A click on a cell puts its address at the caret (selected text there is replaced), a drag a range's (`B2:D5`), a letter a column's (`F:F`), a number a row's (`3:3`). The next click before anything is typed replaces that reference, Shift and a click extend it from its first corner; the field keeps the focus and the active cell stays. Elsewhere a click commits the field as before. A finger does not point.

In a field typing opened over a cell (Enter mode; F2 switches it), an arrow with nothing pointed does what Excel does there: where a reference goes in, it puts the address of the cell beside the edited one (Shift and the arrow a range from the edited cell, `=C3:D3`); after an operand of a formula that reads in the book's locale (`=5`, `=A1`, `=SUM(A1)`), and after any value, it enters the text and moves the selection that way, as Enter does; where the formula does not read (`=A1)`, `=1e`) or inside a string, a quoted sheet's name or a table's brackets, it does nothing. Once a reference is pointed, the keys move it as they move the selection on the sheet: an arrow from its first cell, Shift and an arrow its far corner, Command and an arrow on a Mac (Control elsewhere) to the edge of the data by the source's `end` (with Shift extending there), Home to the row's first column, Page Down and Page Up a screen down and up. Once something is typed after a reference, the next arrow starts from the edited cell again. In a field F2 opened and in the formula bar the arrows move the caret.

## A formula that does not read

A formula the source calls `invalid` (Enter, Tab, Control+Return, the formula bar's Enter or a click away) brings a question with the source's reason: Fix, the default, opens the field again at the error; Restore, or Esc, leaves the cell as it was.

When the outcome carries `correction`, the formula Excel offers for a formula left open (`=(1+2` corrected to `=(1+2)`), the sheet shows Excel's window instead: "We found a typo in your formula and tried to correct it to: =(1+2). Would you like to accept this correction?" with Yes, the default, and No (`SheetMessages`' `typoFound`, `validationYes`, `validationNo`). Yes sends the same call again with the correction, an edit or a Control+Return's fill, and the selection goes where the key that entered it puts it (down after Enter, right after Tab, the range kept after Control+Return). No, or Esc, brings Excel's message "This formula is missing an opening or closing parenthesis." with OK (`missingParenthesis`), after which the field opens again with the text typed, the caret at the error. The sheet adds nothing to a formula itself: the correction is only the source's.

## What the source's optional methods turn on

The source's optional methods turn on the rest:

- Control+Return (Cmd+Return on a Mac), in the cell's field or the formula bar, enters the text into every cell of a selection through `fill(sheet, range, text, active)`, as Excel enters it from the active cell; the selection stays, and the text goes even when it did not change. Without `fill`, or on one cell, the text goes to the active cell through `edit` when it changed, as Enter sends it, and the cursor does not move.
- A paste on the grid goes to `paste(sheet, cell, text)` in one call at the selection's top-left cell, and the source reads the text as Excel pastes it. Without `paste` the component lays the text out as Excel pastes it (xlcalc's `parseClipboard`: a row at each line break, a field at each tab, a line break at the end one more row of empty fields) and sends the fields to `edit`; a rectangle that would leave the sheet is refused. A paste of the sheet's own copy goes through `structure` instead ([clipboard.md](clipboard.md)).
- Control+Z and Control+Y (Cmd on a Mac; Control+Shift+Z redoes too; any keyboard layout) on the grid call `undo()` and `redo()`, one call at a time. In the field they undo the typing. Without the methods there is no undo.
- A successful outcome's `select` is where the selection goes after a paste, a fill, an undo or an edit; on another sheet the component goes there as the address box does. A selection moved while the call was out stays where it was moved.
- A window's `editText` is the text a cell opens with, and the formula bar shows the same text, as Excel's does; a cell left out of it opens empty, and a `CellUpdate` with `editText` changes it. Without it the component writes the text from the formula and the value.
- Columns are drawn in xlcalc's geometry, the one it lays out `####`, overflow and wrapped text by (Excel for Mac): a stored width is `columnWidthToPoints(width)` points, × 4/3 px rounded to a whole pixel, and a column whose width is `null` in a window is 53 points, 71 px. A width dragged with the mouse goes back through `columnWidthFromPoints`, in whole points, and the cells it shows otherwise come back in the `changes` of `layout`'s outcome and are drawn (a date past its `####`), as an edit's are; the Column Width box counts characters of the stored width, as Excel's.
- A window's `appearance` is what conditional formats give its cells, as xlcalc computes it (`ConditionalAppearance`), and the grid draws it as Excel does: a rule's style over the cell's own (font colour, bold, italic, underline, strike, fill, borders; a fill hides grid lines as the cell's own does), a colour scale's colour as a fill, a data bar from its `start` to its `end` (from the right with `rightToLeft`; a gradient from the bar's colour at its axis toward white, or solid; a border; the axis dashed across the cell), an icon of any of the 20 sets at the cell's left before its text, and with `hideValue` the bar or the icon alone (an icon alone stands by the cell's own alignment, at the left for General). A cell's `appearance` takes the place of its `bars` entry; `bars` alone is drawn as a gradient bar from the cell's left edge. The text and colour a rule's number format gives are the window's `text` and `colors`.
- `subscribe(listener)` tells the sheet of changes made elsewhere. `changed` patches the loaded windows; `reload` reads the sheets and windows again and applies a display filter again; `repaint` with ranges reads again only the windows that meet them on the sheet shown, in place: each stays drawn until its answer comes, and the sizes, the scroll and the display filter's rows stay as they are (a row an edit took out of the filter stays until the filter is applied again, as Excel keeps it until Reapply). Repaints of one task are read together, and one that comes within 80 ms of the last aborts its request and asks once the stream stops; a `changed` that comes while a window is read goes over its answer, and a `changed` of cells off screen draws nothing. In the example after this list `memorySource` plays another tab's edit and a server's refusal.
- Control+B, Control+I and Control+U (Cmd on a Mac, any keyboard layout) on the grid toggle bold, italic and underline over the selection by the active cell's look, through `format`, as the toolbar's buttons do. Without `format` the browser keeps the keys.
- A click on a column letter or a row number selects the whole line to Excel's last row or column, so the `state` event and the address box give `B:D`, `2:5` or `1:1048576` for the whole sheet, and formatting goes on the lines.

```vue
<script setup lang="ts">
import { SheetView, memorySource } from 'xlcalc-sheet'
import 'xlcalc-sheet/style.css'

const source = memorySource({ sheets: [{ name: 'Sheet1', cells: { A1: 'Visitors', B1: 120 } }] })
// Another tab writes B1: the sheet gets a `changed` event and redraws the cell
const remoteEdit = () => source.setCells('Sheet1', { B1: String(Math.round(Math.random() * 1000)) })
// The next edit is refused, as a server would refuse a locked cell
const refuseNext = () => source.failNext({ ok: false, reason: 'refused', message: 'This cell is locked on the server' })
</script>

<template>
  <button @click="remoteEdit">Edit B1 elsewhere</button>
  <button @click="refuseNext">Refuse the next edit</button>
  <SheetView :source="source" style="height: 300px" />
</template>
```
