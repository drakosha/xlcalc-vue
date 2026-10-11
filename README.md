# xlcalc-sheet

Vue 3 sheet component for [xlcalc](https://github.com/drakosha/xlcalc) workbooks: a windowed grid with the book's styles, formulas and editing, which talks to its data only through a `SheetSource` (`src/source.ts`). MIT.

```vue
<script setup lang="ts">
import { SheetView, memorySource, ruMessages } from 'xlcalc-sheet'
import 'xlcalc-sheet/style.css'

const source = memorySource({ sheets: [{ name: 'Sheet1', cells: { A1: 'Hello', B1: 42 } }] })
</script>

<template>
  <SheetView :source="source" :messages="ruMessages" storage-key="my-book" style="height: 480px" @state="console.log($event)" />
</template>
```

The runtime depends only on `vue`: xlcalc's engine-free subpaths `xlcalc/a1` and `xlcalc/text` (addresses, the references of formula text, a paste's rectangle) are bundled in, and a book in ru-RU loads ru-RU's formula text as a chunk of its own. Colours and fonts are `--xs-*` CSS variables. `npm run dev:example` serves the example with every kind of formatting, `npm run screens` takes its screenshots, `docs/host-parity.md` maps the component onto the the app's grid.

## Keys and the source

Editing takes Excel's keys: typing over a cell, F2 or a double click opens the field, Enter, Tab and their Shift commit and move, Esc cancels, Alt+Enter breaks the line, Delete and Backspace clear the selection. Without `edit` the sheet is read-only. The source's optional methods turn on the rest:

- Control+Return (Cmd+Return on a Mac), in the cell's field or the formula bar, enters the text into every cell of a selection through `fill(sheet, range, text, active)`, as Excel enters it from the active cell; the selection stays, and the text goes even when it did not change. Without `fill`, or on one cell, the text goes to the active cell through `edit` and the cursor does not move.
- A paste on the grid goes to `paste(sheet, cell, text)` in one call at the selection's top-left cell, and the source reads the text as Excel pastes it. Without `paste` the component lays the tab-separated text out itself and sends the fields to `edit`.
- Control+Z and Control+Y (Cmd on a Mac; Control+Shift+Z redoes too; any keyboard layout) on the grid call `undo()` and `redo()`, one call at a time. In the field they undo the typing. Without the methods there is no undo.
- A successful outcome's `select` is where the selection goes after a paste, a fill, an undo or an edit; on another sheet the component goes there as the address box does. A selection moved while the call was out stays where it was moved.
- A window's `editText` is the text a cell opens with, and the formula bar shows the same text, as Excel's does; a cell left out of it opens empty, and a `CellUpdate` with `editText` changes it. Without it the component writes the text from the formula and the value.
- A column whose width is `null` in a window is 8.83203125, xlcalc's default (Excel for Mac, Calibri 11), which is 62 px.

## A book computed in the browser: `xlcalc-sheet/local`

`localSource(book)` is a `SheetSource` over an xlcalc `BookInput`. It needs `xlcalc` 0.19 or later, an optional peer of the package that the host installs. The book computes in a Web Worker (`xlcalc/worker`), or in the page's thread with `calc: 'inline'`. The book's locale module is loaded when the source opens (ru-RU from `xlcalc/locales/ru-RU`); a locale xlcalc has no module for fails the opening with the library's message.

The session does the work as Excel does it, and the source passes it on with every optional method of `SheetSource`. A window is one `getViewport` call: formulas as the book's locale shows them, the text each cell opens with (`editText`), styles, widths, heights, merged areas and tables. Typed text is read in the book's locale, several cells in one step of the undo (`=СУММ(A1:A2)`, `1 234,5`, `15.03.2026` in ru-RU), and a column a typed number does not fit widens; Delete over a selection clears the content and the hyperlinks, as Excel's Delete. `paste` takes clipboard text by Excel's rules for a paste (an apostrophe stays in the text, a percent cell has no percent editor), `fill` is Control+Return, `undo` and `redo` are the session's, whole. A value a stop rule of data validation does not take is refused and taken back. Structure operations, widths and heights, formatting over each cell's look (`format`, as Excel's Bold over a selection), `explain`, totals, search, the display filter and sort complete it. `openXlsx` reads a file into a book and `saveXlsx` writes the source's book with its values.

```ts
import { SheetView } from 'xlcalc-sheet'
import { localSource, openXlsx, saveXlsx } from 'xlcalc-sheet/local'

const { book, issues } = await openXlsx(file) // a File from an input or a drop; issues: what the book cannot hold
const source = localSource(book)               // <SheetView :source="source" :locale="source.locale" />
const { data } = await saveXlsx(source)        // the xlsx bytes, formulas with their values
```

The pool starts its thread as `new Worker(new URL('./thread.js', import.meta.url))`, which Vite and other bundlers follow; where a bundler serves the thread elsewhere, pass `pool: { spawn }`. Besides `SheetSource` the source has `ready`, `locale`, `history` (the steps `undo` and `redo` can take, for a toolbar), `book()`, `snapshot()` (the book and its values taken at one moment, what `writeXlsx` takes; it replaces `result()` of 0.2), `freeze()` and `close()`. It tells its subscribers of every change it makes: `changed` for edits and widths, `reload` for structure, styles, lines, frozen panes and their undo. `npm run check:local` opens the example's `local.html` (the synthetic book of 3,400 rows) in Chromium, measures it and takes screenshots; the page also opens an xlsx file from its field and saves the book under the file's name.
