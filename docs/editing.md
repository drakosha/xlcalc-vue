# Cell editing and keyboard shortcuts

Editing uses Excel's keys:

- typing over a cell, F2 or a double click starts editing the cell;
- Enter, Tab, Shift+Enter and Shift+Tab commit the entry and move the active cell. Inside a selection of several cells, the selection stays and the active cell moves within it as in Excel: Enter goes down a column and on to the top of the next one, Tab goes along a row and on to the start of the next one, wrapping around at the end;
- Esc cancels the edit, and Alt+Enter inserts a line break;
- Ctrl+Shift+U (or the arrow at the right end of the formula bar) expands the formula bar to five lines and collapses it back to one, as in Excel; while it is collapsed, a multi-line cell shows its first line and the grid does not move;
- Delete and Backspace clear the selection;
- Ctrl+` shows the formulas instead of the values, as Excel's Show Formulas does ([view.md](view.md#show-formulas));
- Ctrl+Page Down and Ctrl+Page Up go to the next and the previous visible sheet;
- F5 and Ctrl+G (Control+G on a Mac as well) open Excel's Go To dialog ([go-to-and-symbol.md](go-to-and-symbol.md));
- Ctrl+F3 opens the Name Manager, and Ctrl+L on a Mac Define Name, with a source that has defined names ([names.md](names.md));
- F7 opens Spelling when the source has `spelling` ([Spelling](#spelling)), and Ctrl+T or Ctrl+L (Control+T on a Mac) opens Create Table when it has `table` ([tables.md](tables.md));
- Ctrl+arrow goes where the source's `end(sheet, cell, direction)` points, as Excel's Range.End, and Shift+Ctrl+arrow extends the selection there. Without `end`, or under a display filter, it goes to the edge of the data known to the sheet.

Without `edit` the sheet is read-only.

A workbook calculated in the browser to try the keys on: type into B2 or C3 and D4 updates, F2 edits a cell in place, Ctrl+Enter over B2:C3 enters the same text into all four cells, and `=(1+2` brings up Excel's typo dialog.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

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

The formula bar is an input that always contains the active cell's text, the same text the cell editor opens with (a formula with its `=`), as in Excel; while the cell editor is open, it shows the draft. Clicking in it places the caret at the click and starts editing there, as F2 does; dragging starts editing with the dragged-over text selected; and clicking it while the cell editor is open moves the draft to the formula bar with the caret at the click. Focusing the input also starts editing, so anything typed into it always goes to the cell.

A formula's references are colored like their range borders on the grid, including those of a draft as it is typed. Outside editing, Ctrl+click on a reference (Cmd+click on a Mac) selects its range, on another sheet as well, like Go to Definition in a code editor, and the input's tooltip says so (`SheetMessages`' `goRefHint`). While Ctrl or Cmd is held, the reference under the pointer is underlined in its color and the pointer becomes a hand, whichever came first, the key or the mouse; both go away when the key is released, the pointer leaves the reference, or the window loses focus. On a touch screen, a long press on a reference (half a second, like the long press that opens the context menus; moving the finger cancels it) goes to its range instead, starts no edit and suppresses the browser's own menu and text selection, while a tap places the caret and edits as before; the input has no tooltip there. On a read-only sheet (without `edit`), on a locked cell, or with nothing selected, the input is read-only: its text can be selected and copied, keys do not start editing, and Ctrl+click or a long press on a reference still goes to it.

To the left of the input are Excel's Cancel (✕, `SheetMessages`' `cancel`) and Enter (✓, `barEnter`) buttons, enabled only while the cell editor or the formula bar is being edited. Cancel works like Esc; Enter enters the draft and keeps the active cell in place, like Ctrl+Enter on a single cell, and a formula that cannot be parsed brings up the same prompt as with Enter. Pressing either button keeps focus in the input until the button acts. They are followed by fx, Excel's Insert Function ([autocomplete.md](autocomplete.md#insert-function)). The reference colors are a layer over the input's transparent text, drawn with the CSS Custom Highlight API, so the layer's text lines up with the input's to the pixel; a browser without the API gets colored spans, which may be off by a fraction of a pixel.

## Point mode

While a formula is being typed in the cell editor or the formula bar, the mouse inserts references as in Excel's Point mode, wherever xlcalc's `refAt` (measured in Excel) allows a reference: after the formula's first character (`=`, or a leading `+` or `-`), an operator, `(`, the workbook's list separator (`;` in ru-RU), `:`, a sheet's `!`, a decimal point, or a space after an operand (Excel's intersection operator). Clicking a cell inserts its address at the caret (replacing any text selected there), dragging inserts a range (`B2:D5`), clicking a column letter inserts the column (`F:F`), and clicking a row number inserts the row (`3:3`). Another click before anything is typed replaces that reference, and Shift+click extends it from its first corner; the input keeps focus and the active cell stays in place. Elsewhere a click commits the entry as before. Touch input does not insert references.

In a cell editor opened by typing over a cell (Enter mode; F2 switches modes), an arrow key with no reference inserted yet does what Excel does in that position: where a reference can go, it inserts the address of the adjacent cell (Shift+arrow inserts a range from the edited cell, `=C3:D3`); after an operand of a formula that parses in the workbook's locale (`=5`, `=A1`, `=SUM(A1)`), and after any value, it enters the text and moves the selection in that direction, as Enter does; where the formula does not parse (`=A1)`, `=1e`) or inside a string, a quoted sheet name or a table's brackets, it does nothing. Once a reference is inserted, the keys move it as they move the selection on the sheet: an arrow moves it from its first cell, Shift+arrow moves its far corner, Cmd+arrow on a Mac (Ctrl+arrow elsewhere) moves it to the edge of the data by the source's `end` (Shift extends it there), Home moves it to the row's first column, and Page Down and Page Up move it a screen down or up. Once anything is typed after a reference, the next arrow starts from the edited cell again. In an editor opened with F2 and in the formula bar, the arrow keys move the caret.

## Formulas that cannot be parsed

When the source rejects a formula as `invalid` (on Enter, Tab, Ctrl+Enter, the formula bar's Enter button or a click elsewhere), the sheet asks what to do and shows the source's reason: Fix, the default, reopens the editor at the error; Restore, or Esc, leaves the cell unchanged.

When the outcome carries `correction`, the formula Excel suggests for an unclosed formula (`=(1+2` corrected to `=(1+2)`), the sheet shows Excel's dialog instead: "We found a typo in your formula and tried to correct it to: =(1+2). Would you like to accept this correction?" with Yes, the default, and No (`SheetMessages`' `typoFound`, `validationYes`, `validationNo`). Yes sends the same call again with the correction, an edit or a Ctrl+Enter fill, and the selection moves as the key that entered it moves it (down after Enter, right after Tab, the range kept after Ctrl+Enter). No, or Esc, shows Excel's message "This formula is missing an opening or closing parenthesis." with OK (`missingParenthesis`), after which the editor reopens with the typed text and the caret at the error. The sheet never corrects a formula on its own: the correction comes only from the source.

## What the optional source methods enable

The source's other optional methods enable the following:

- Ctrl+Enter (Cmd+Return on a Mac), in the cell editor or the formula bar, enters the text into every cell of the selection through `fill(sheet, range, text, active)`, as Excel enters it from the active cell; the selection stays, and the text is sent even if it did not change. Without `fill`, or with a single cell, the text goes to the active cell through `edit` if it changed, as with Enter, and the active cell does not move.
- Ctrl+D and Ctrl+R (Cmd or Control on a Mac, as Excel for Mac takes either; with any keyboard layout) are Excel's Fill Down and Fill Right through `structure`, as a single `copyRange`: the selection's top row is copied over its other rows (its left column over its other columns) with its values, formulas and formatting, relative references shifting as Excel's fill shifts them and absolute ones kept. A selection one row tall takes the row above it, one column wide the column to its left, and row 1 alone (column A alone) is left as it is. Entire columns copy their first row down the whole sheet, and entire rows their first column across it. On a sheet the workbook itself filters, xlcalc's `copyRange` fills only the visible rows, as Excel's Fill Down does. Under the sheet's display filter, whose hidden rows the source does not know, nothing is filled, as nothing is pasted there from the sheet's own copy; a locked cell among the filled ones is reported instead, as is what the source refuses (xlcalc refuses a copy that would split a merged area or part of an array, as Excel does). The selection stays as it was. On an editable sheet the keys do not reach the browser (its bookmark and reload), even where nothing can be filled; on a read-only sheet they do. A host's ribbon runs the same commands through `fillDown()` and `fillRight()` ([api.md](api.md#sheetviewapi)); Excel's cell context menu has no Fill commands, and neither does the sheet's.
- A paste on the grid goes to `paste(sheet, cell, text)` as a single call at the selection's top-left cell, and the source parses the text as Excel pastes it. Without `paste`, the component splits the text as Excel pastes it (xlcalc's `parseClipboard`: a row at each line break, a field at each tab, and one more row of empty fields after a final line break) and sends the fields to `edit`; a range that would extend past the edge of the sheet is rejected. Pasting the sheet's own copy goes through `structure` instead ([clipboard.md](clipboard.md)).
- Ctrl+Z and Ctrl+Y (Cmd on a Mac; Ctrl+Shift+Z also redoes; with any keyboard layout) on the grid call `undo()` and `redo()`, one call at a time. In an editor they undo the typing. Without these methods there is no undo.
- A successful outcome's `select` sets where the selection goes after a paste, a fill, an undo or an edit; if it is on another sheet, the component switches to it as the Name Box does. If the selection was moved while the call was pending, it stays where it was moved.
- A window's `editText` is the text a cell opens with for editing, and the formula bar shows the same text, as Excel's does; a cell missing from it opens empty, and a `CellUpdate` with `editText` changes it. Without it the component produces the text itself: the formula with its `=`, otherwise `entryText` from `xlcalc/entry` for the value in the workbook's locale, according to the cell's code in `numberFormats` (General if there is none), as Excel's formula bar displays it: a number in full with the locale's decimal separator, a percentage cell's value as a percentage, a date or time cell's value in the locale's short date and long time formats, booleans and errors in the locale's words, and text with Excel's apostrophe where typing it in that locale would be parsed as something else (`'007`; `'1,5` in ru-RU, `1,5` in en-US). `xlcalc/entry` is a separate chunk, loaded when the first cell needs its text (the active cell, for the formula bar); until it has loaded, an editor opens with the displayed text right away, and the formula bar switches to the produced text once it is available.
- Columns are drawn with xlcalc's geometry, the one it uses to lay out `####`, overflow and wrapped text (Excel for Mac's): a stored width is `columnWidthToPoints(width)` points, times 4/3 px rounded to a whole pixel, and a column whose width is `null` in a window is 53 points (71 px). A width dragged with the mouse is converted back through `columnWidthFromPoints`, in whole points, and the cells whose display changes with it come back in the `changes` of `layout`'s outcome and are redrawn, as an edit's are (a date that no longer shows `####`). The Column Width box counts characters of the stored width as Excel for Mac's does (`columnWidthToChars` and `columnWidthFromChars`): the default column is 8, and typing 8 gives the default back.
- A window's `appearance` is the conditional formatting of its cells as xlcalc calculates it (`ConditionalAppearance`), and the grid draws it as Excel does: a rule's style over the cell's own (font color, bold, italic, underline, strikethrough, fill, borders; a fill hides gridlines as the cell's own fill does), a color scale's color as a fill, a data bar from its `start` to its `end` (from the right with `rightToLeft`; a gradient from the bar's color at its axis toward white, or a solid fill; a border; a dashed axis across the cell), an icon from any of the 20 icon sets at the left of the cell before its text, and with `hideValue` only the bar or the icon (an icon shown alone follows the cell's own alignment, left for General). The text and color produced by a rule's number format are the window's `text` and `colors`.
- `subscribe(listener)` notifies the sheet of changes made elsewhere. `changed` patches the loaded windows; `reload` reads the sheets and windows again and reapplies a display filter; `repaint` with ranges re-reads in place only the windows of the displayed sheet that overlap the ranges: each window stays drawn until its response arrives, and the sizes, the scroll position and the display filter's rows stay as they are (a row that no longer matches the filter after an edit remains until the filter is reapplied, as Excel keeps it until Reapply). Repaints from one task are read together, and a repaint that arrives within 80 ms of the previous one aborts its request and reads once the burst ends; a `changed` that arrives while a window is being read is applied over its response, and a `changed` for cells off screen draws nothing. In the example after this list, `memorySource` simulates an edit from another tab and a rejection by a server.
- Ctrl+B, Ctrl+I and Ctrl+U (Cmd on a Mac, with any keyboard layout) on the grid toggle bold, italic and underline over the selection according to the active cell's formatting, through `format`, as the toolbar buttons do. Ctrl+Shift+% (Control on a Mac too, as in Excel for Mac, since Cmd+Shift+5 is the system's screenshot key; the key of 5 on any layout) applies Percent Style, Excel's `0%` (the built-in format 9), as the toolbar's button does. Without `format` the browser handles these keys.
- Clicking a column letter or a row number selects the entire column or row up to Excel's last row or column, so the `state` event and the Name Box report `B:D`, `2:5`, or `1:1048576` for the entire sheet, and formatting applies to entire rows or columns.

```vue
<script setup lang="ts">
import { SheetView, memorySource } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({ sheets: [{ name: 'Sheet1', cells: { A1: 'Visitors', B1: 120 } }] })
// Another tab writes B1: the sheet gets a `changed` event and redraws the cell
const remoteEdit = () => source.setCells('Sheet1', { B1: String(Math.round(Math.random() * 1000)) })
// The next edit is rejected, as a server would reject an edit of a locked cell
const refuseNext = () => source.failNext({ ok: false, reason: 'refused', message: 'This cell is locked on the server' })
</script>

<template>
  <button @click="remoteEdit">Edit B1 elsewhere</button>
  <button @click="refuseNext">Refuse the next edit</button>
  <SheetView :source="source" style="height: 300px" />
</template>
```

## Spelling

If the source has `spelling(req)`, F7 (also on a Mac, where Excel uses it too) and the sheet's `checkSpelling()` method open Excel's Spelling dialog. The check starts at the active cell, goes by rows through the entire sheet, or only through the selection when more than one cell is selected, and wraps around to the cells before the active cell; formulas, words in uppercase and words with digits are not checked, as Excel's defaults. For each misspelt word the sheet selects its cell, and the dialog shows the cell's text with the word marked under Not in Dictionary and the source's corrections under Suggestions, the first one chosen (Up and Down choose another, a double click is Change):

- Ignore Once goes on to the next word, in the same cell first (`after` is the cell, `from` the position after the word);
- Ignore All skips the word for the rest of this check (`ignore`);
- Add to Dictionary, present when the source has `learnWord`, makes the word the user's own and goes on;
- Change writes the cell's text with the word replaced by the chosen suggestion through `edit`, as typed text (a text that opened with an apostrophe keeps it), and goes on after the new word;
- Change All does the same, and replaces the word in each cell it is found in afterwards without stopping there, as Excel goes on;
- Cancel or Esc ends the check.

When the check comes back to where it started, or there is nothing to correct, Excel's message says that the check is complete. On a read-only sheet Change and Change All are disabled; a locked cell is not changed. The editor keeps `spellcheck="false"`: Excel does not check spelling as you type. `memorySource` and `./local` check spelling with the host's `SpellChecker` (a dictionary such as hunspell's), given as their `spelling` option.

```vue
<script setup lang="ts">
import { SheetView, memorySource } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

// A tiny word list for the example; a host passes a real dictionary
const words = new Set(['the', 'quick', 'brown', 'fox', 'jumps', 'over', 'lazy', 'dog'])
const own = new Set<string>()
const source = memorySource(
  { sheets: [{ name: 'Notes', cells: { A1: 'teh quick brown fox', A2: 'jumps ovr the lazy dog', A3: 'teh end' } }] },
  {
    spelling: {
      check: (w) => words.has(w.toLowerCase()) || own.has(w.toLowerCase()),
      suggest: (w) => ({ teh: ['the', 'ten'], ovr: ['over'] })[w.toLowerCase()] ?? [],
      add: (w) => void own.add(w.toLowerCase()),
    },
  },
)
</script>

<template>
  <!-- click the grid and press F7 -->
  <SheetView :source="source" style="height: 300px" />
</template>
```

## AutoSum

With a source that has `autoSum`, Alt+= (Option+= on a Mac, by the key's position on any layout) is Excel's AutoSum, and so is the sheet's `autoSum()` method, which a host's ribbon button calls. The source guesses the range as Excel does (`./local` through xlcalc's `getAutoSum`: the run of numbers above the cell, else to its left, up to a subtotal), and the active cell's edit opens with `=SUM(range)` in the source's locale and the range selected: Enter enters it, typing replaces the range and the arrows point at another, as in Point mode. Without numbers the edit opens with `=SUM()` and the caret between the parentheses. Excel's AutoSum list has the same guess under other functions: `autoSum('average')`, `'count'` (Count Numbers, `COUNT`), `'max'` and `'min'` open the edit with `=AVERAGE(range)` and so on, for a source that lists them in `autoSumFunctions` and builds them when `autoSum` gets the function as its fourth argument; `./local` builds all four over xlcalc's range and translates them into the workbook's locale with xlcalc's `toLocal` (`=СРЗНАЧ(A4;A2)` in ru-RU). The toolbar's Σ and its list call these ([toolbar.md](toolbar.md)). Excel's AutoSum over a selection of several cells, which enters the sums at once, is not offered: the sheet sums at the active cell. Excel also gives the sum the number format of what it sums (a sum of percentages shows as a percentage); xlcalc does not yet.

