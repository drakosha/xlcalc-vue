# Go To, Go To Special and Symbol

## Go To

F5 and Ctrl+G on the grid open Excel's Go To dialog (Home > Find & Select > Go To); on a Mac it is Control+G, since Cmd+G finds again in the browser. The host opens it with the sheet's `goTo()` method ([api.md](api.md#sheetviewapi)). The dialog has the Go to list, the Reference box, Special… (with a source that has `specialCells`, see [Go To Special](#go-to-special)), OK and Cancel.

The list shows the last four locations that jumps from Go To and from hyperlinks into the workbook have left, newest first, each once, written as Excel lists them: absolute and with the sheet's name (`Data!$B$2:$D$5`, `'Sheet 2'!$C:$C`). The Reference box opens with the first of them, so F5 and Enter go back to where the last jump came from, and F5 and Enter again return. A click on a location puts it into the Reference box, the up and down arrows step through the list from anywhere in the dialog, and a double click goes there. The locations last as long as the workbook does in the component. With a source that has `names`, the defined names that refer to ranges follow the locations, alphabetically, as the Name Box lists them, and OK on a name goes to its range ([names.md](names.md)).

OK (or Enter) goes to the reference as the Name Box does: a cell, a range, entire columns or rows, another sheet (`'Sheet 2'!B7`), with the flash and a Back history entry. Text that is not such a reference brings up Excel's alert, `Reference isn’t valid.`, with OK alone; afterwards the dialog stays with the text selected. Esc, Cancel or a click outside the dialog closes it and returns to the grid. The dialog works on a read-only sheet too. On a phone it takes the full screen, and on a touch screen its inputs and items are 44 px. Its texts are `SheetMessages`' `goToTitle`, `goToList`, `goToReference`, `goToBad` and `goToSpecial`, with `ok` and `cancel`.

## Go To Special

Excel's Go To Special (Home > Find & Select > Go To Special) opens from the Go To dialog's Special…, which closes Go To, and through the sheet's `goToSpecial()` for a ribbon. It needs the source's `specialCells` (`./local`: xlcalc 0.66.0 or later) and no display filter, whose rows on screen are not the sheet's; it works on a read-only sheet too.

The dialog lays out Excel's options in its two columns, with the options the component can find. Constants and Formulas, with the kinds of values under them (Numbers, Text, Logicals, Errors, all checked at first; a formula counts by the value it answers), Blanks, Current region; Last cell, Visible cells only, Conditional formats and Data validation, with All or Same under them. It opens with the choice last made, Constants at first. OK, or Enter, closes it and selects what it found:

- Constants, Formulas, Blanks and Visible cells only are the source's `specialCells` over each area of the selection; a single selected cell stands for the sheet's used range, as in Excel, so Blanks finds the blanks among the data (the used range is the library's, measured in Excel: it reaches rows and columns of their own height, width or style). `./local` gives xlcalc's `getSpecialCells`.
- Current region is the active cell's region, read from the windows as Create Table reads it ([tables.md](tables.md)).
- Last cell is the used range's last cell, whatever the selection (`specialCells` with `{ type: 'lastCell' }`); `./local` takes the last row and column of the cells, spills included, and of the rows and columns of their own that the used range reaches.
- Conditional formats and Data validation read the sheet's rules (the source's `rules`; without it both are disabled). All takes the cells of every rule, within the selection unless it is a single cell; Same ignores the selection and takes, for conditional formats, the cells of the first rule (by priority) over the active cell, and for data validation the cells of every rule with the same settings as the active cell's. Both are as Excel for Mac selected them on a probe workbook.

The cells found are selected as Excel selects them, several ranges at once (see below). The active cell stays where it is when one of the ranges holds it, otherwise it goes to the first range's top-left cell, as Excel for Mac does. When nothing is found, Excel's alert says `No cells were found.` and the selection stays as it was. A read the source fails goes to the message line.

Excel's other options are left out, since no source method tells them: Notes (Excel's Comments), which needs every note of the sheet, and Excel's dialog opens with it; Current array, which needs the extent of the active cell's array; Objects, which the component has none of; Row differences and Column differences, which compare values by Excel's rules; Precedents and Dependents, which need every level of references on the sheet (`dependents` gives direct dependents only, a page at a time). The dialog's texts are `specialTitle`, `specialSelect`, `specialConstants`, `specialFormulas`, `specialNumbers`, `specialText`, `specialLogicals`, `specialErrors`, `specialBlanks`, `specialRegion`, `specialLastCell`, `specialVisible`, `specialConditional`, `specialValidation`, `specialAll`, `specialSame` and `specialNone`, with `ok` and `cancel`; the English and Russian ones are Excel for Mac's.

### A selection of several ranges

Go To Special's ranges make one selection, as in Excel: the range of the active cell is drawn as the selection is, the others with the selection's tint (`--xs-sel-tint`, cells with the class `xs-area`), and the Name Box shows the active cell. Enter and Tab move the active cell through its range and on to the next one (Shift: back to the previous one), as Excel does; any other move of the active cell, a click or a jump, ends the selection of several ranges. Delete clears the cells of every range, Clear Formats and Clear All clear every range, and formatting (the toolbar, Ctrl+B, Format Cells, the sheet's `format` without a range) formats every range, each as a single undo step. Copy, Cut, Paste and Paste Special are refused with Excel's message, `This action won’t work on multiple selections.` (`multipleSelections`). The other commands act on the active cell's range alone: in Excel some of them act on every range (Ctrl+Enter enters the text into each, Fill and Insert work per range), which the source's methods, made for one range, do not carry.

## Symbol

The sheet's `insertSymbol()` method opens Excel's Symbol dialog (Insert > Symbol) on an editable sheet with an active cell; Excel has no key for it, so the host gives it a button. The dialog has two tabs.

Symbols: the Subset box, the grid of characters, the recently used symbols and the Character code box. The grid runs through the Unicode blocks of Excel's Subset list that a text font usually covers (Basic Latin, Latin-1 Supplement, the Latin extensions, Greek and Coptic, Cyrillic, General Punctuation, Currency Symbols, Letterlike Symbols, Number Forms, Arrows, Mathematical Operators, Box Drawing, Geometric Shapes, Miscellaneous Symbols, Dingbats and the others in `SymbolSubset`), with every assigned character in code order and without control characters. The Subset box shows the selected character's block, and choosing a block goes to its first character. The grid is 16 characters wide, as Excel's, or as many as fit on a narrow screen; the arrow keys move by a character and by a row, Home and End to the ends of the row (with Ctrl, of the grid), and Page Up and Page Down by the rows in view. The Character code box shows the selected character's Unicode code in hex (`00A9`) and selects the character whose code is typed there, with or without `U+`; a code outside the grid's blocks (`1F600`) is inserted too. The recently used symbols are a row of 16, newest first: Office's row (`€ £ ¥ © ® ™ ± ≠ ≤ ≥ ÷ × ∞ µ α β`) until something is inserted, then kept in the browser's storage under `storageKey`, otherwise for the lifetime of the component. The dialog opens on the last character inserted.

Special Characters: Excel's list by name, Em Dash, En Dash, Nonbreaking Hyphen, Optional Hyphen, Em Space, En Space, 1/4 Em Space, Nonbreaking Space, Copyright, Registered, Trademark, Section, Paragraph, Ellipsis, the single and double opening and closing quotes, No-Width Optional Break and No-Width Non Break.

The characters are shown in the active cell's font, Excel's "(normal text)"; the dialog has no font box. Insert, Enter or a double click inserts the selected character of the tab on display. While the cell or the formula bar is being edited, the character goes at the caret in place of the selected text, and the edit stays open while the dialog is shown. Otherwise Insert starts editing the active cell with the character as typing does, so the character replaces the cell's contents, as in Excel; to add to them, start the edit with F2 or a double click first. The dialog stays open for more characters, each after the last, and its Cancel button becomes Close; closing it (Esc, Close, a click outside) returns to the edit with the caret after the characters inserted, where Enter enters them, or to the grid when nothing was inserted. A locked cell gets the sheet's message instead. On a phone the dialog takes the full screen, and on a touch screen its cells, tabs and items are 44 px. Its texts are `SheetMessages`' `symbolTitle`, `symbolTab`, `symbolSubsetLabel`, `symbolSubset`, `symbolRecent`, `symbolCode`, `symbolCodeFrom`, `symbolCharacter`, `symbolSpecial` and `symbolInsert`, with `cancel` and `close`.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { SheetView, memorySource, type SheetViewApi } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({ sheets: [{ name: 'Notes', cells: { A1: 'Price', B1: 'EUR 120' } }] })
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <!-- mousedown.prevent keeps an open edit open: the symbol goes at its caret -->
  <button @mousedown.prevent @click="sheet?.insertSymbol()">Symbol…</button>
  <button @click="sheet?.goTo()">Go To…</button>
  <SheetView ref="sheet" :source="source" style="height: 360px" />
</template>
```
