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
- A paste on the grid goes to `paste(sheet, cell, text)` in one call at the selection's top-left cell, and the source reads the text as Excel pastes it. Without `paste` the component lays the text out as Excel pastes it (xlcalc's `parseClipboard`: a row at each line break, a field at each tab, a line break at the end one more row of empty fields) and sends the fields to `edit`; a rectangle that would leave the sheet is refused.
- Control+Z and Control+Y (Cmd on a Mac; Control+Shift+Z redoes too; any keyboard layout) on the grid call `undo()` and `redo()`, one call at a time. In the field they undo the typing. Without the methods there is no undo.
- A successful outcome's `select` is where the selection goes after a paste, a fill, an undo or an edit; on another sheet the component goes there as the address box does. A selection moved while the call was out stays where it was moved.
- A window's `editText` is the text a cell opens with, and the formula bar shows the same text, as Excel's does; a cell left out of it opens empty, and a `CellUpdate` with `editText` changes it. Without it the component writes the text from the formula and the value.
- A column whose width is `null` in a window is 8.83203125, xlcalc's default (Excel for Mac, Calibri 11), which is 62 px.
- A window's `appearance` is what conditional formats give its cells, as xlcalc computes it (`ConditionalAppearance`), and the grid draws it as Excel does: a rule's style over the cell's own (font colour, bold, italic, underline, strike, fill, borders; a fill hides grid lines as the cell's own does), a colour scale's colour as a fill, a data bar from its `start` to its `end` (from the right with `rightToLeft`; a gradient from the bar's colour at its axis toward white, or solid; a border; the axis dashed across the cell), an icon of any of the 20 sets at the cell's left before its text, and with `hideValue` the bar or the icon alone (an icon alone stands by the cell's own alignment, at the left for General). A cell's `appearance` takes the place of its `bars` entry; `bars` alone is drawn as a gradient bar from the cell's left edge. The text and colour a rule's number format gives are the window's `text` and `colors`.
- `subscribe(listener)` tells the sheet of changes made elsewhere. `changed` patches the loaded windows; `reload` reads the sheets and windows again and applies a display filter again; `repaint` with ranges reads again only the windows that meet them on the sheet shown, in place: each stays drawn until its answer comes, and the sizes, the scroll and the display filter's rows stay as they are (a row an edit took out of the filter stays until the filter is applied again, as Excel keeps it until Reapply). A second `repaint` aborts the request of the first, and a `changed` that comes while a window is read goes over its answer.
- Control+B, Control+I and Control+U (Cmd on a Mac, any keyboard layout) on the grid toggle bold, italic and underline over the selection by the active cell's look, through `format`, as the toolbar's buttons do. Without `format` the browser keeps the keys.
- A click on a column letter or a row number selects the whole line to Excel's last row or column, so the `state` event and the address box give `B:D`, `2:5` or `1:1048576` for the whole sheet, and formatting goes on the lines.

## Rows, columns and sheets

A right click on a row number or a column letter, or a finger held on it for half a second, opens Excel's menu for the selected lines (a line outside a selection of whole lines is selected first): Insert and Delete as many lines as are selected through `structure(op)`, Row Height or Column Width in Excel's units (points; characters of the default font, the width Excel's box shows) with Excel's limits, Hide, and Unhide on a selection over hidden lines, through `layout(change)`; a size of 0 hides the lines. The same on a tab makes that sheet the active one and offers Insert (before it, under Excel's next name: `Sheet4`, `Лист4` with `ruMessages`), Delete (asked first when the sheet has data), Rename (in place on the tab, a double click too), Move Left and Move Right among the shown tabs, Hide, and Unhide with the list of hidden sheets. A refusal shows the source's message (`./local` gives the library's, in the book's locale), and after a change the sheet reads its sheets and windows again unless the source sent `reload`. The sheet shown after is Excel's: a sheet inserted or unhidden, and for one deleted or hidden the next shown sheet, or the one before it at the end. Without `structure` or `layout` their items are gone, without both the browser's own menu stays, and a read-only sheet has none.

The tabs show with one sheet too and end with "+", which inserts a sheet at the end and shows it (with `structure`); hidden sheets are not on them. A double click on a column's edge fits the column to what it shows through `layout({ columns, fit: true })`, every selected whole column when it is one of them, as Excel does; without `layout` it brings back the book's width. Two properties of `SheetView` give back the the app's ways: `hide-lone-tab` (no tabs for a book of one sheet) and `edge-restores-width` (the double click brings back the book's width).

## The sheet's methods

`SheetView` exposes `undo()`, `redo()`, `format(patches)` and `structure(op)` through a template ref. Each goes through the source as the sheet's own keys go: the outcome is applied, its selection set without the flash and the step back that `focus` makes, a refusal shown in the sheet's message line, and after `format` or `structure` the sheet reads itself again unless the source sent `reload`. `format` formats the active sheet: a patch without `range` goes over the selection as Excel writes it (`B:B`, `2:5`). Each returns the source's outcome, or null when nothing was sent (no such method, a read-only sheet, an undo still out).

## The formatting toolbar: `SheetToolbar`

`SheetToolbar` is Excel's formatting row over the same source: undo and redo, bold, italic, underline, font and fill colours from Excel's palette, borders (all, outside, bottom, none), horizontal alignment, wrap text and a number format list. It formats the selection the sheet's `state` event gives, through `format(sheet, patches)`, as Excel formats a selection: bold goes over each cell's look, outside borders are four patches in one call and so one step of undo. Its buttons are pressed by what the active cell shows, and pressed again they take it off, as Excel's are. It shows only when the source has `format`; undo and redo come with the source's `undo` and `redo`, off while its `history` (the local source's) has no step.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { SheetToolbar, SheetView, ruMessages, type SheetViewApi, type SheetViewState } from 'xlcalc-sheet'

// source: a SheetSource with format, such as localSource(book); formats: the number formats to offer (below)
const state = ref<SheetViewState | null>(null)
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <SheetToolbar :source="source" :sheet="sheet" :state="state" :messages="ruMessages" :number-formats="formats" />
  <SheetView ref="sheet" :source="source" :messages="ruMessages" @state="state = $event" />
</template>
```

The main entry has no engine, so it does not know what code a built-in format shows by in a locale: the host passes the number formats to offer as `numberFormats`, each with what it sets (a code, or a built-in format's number) and the code a cell shows by with it (`{ kind: 'date', format: 14, code: 'dd.mm.yyyy' }` in ru-RU), and the list selects the one whose code the active cell has (General without one, Custom for any other). Without `numberFormats` there is no list; `example/spa-book.ts` builds them with xlcalc's `builtinFormatCode` and `parseEntry`. With `sheet` (the sheet's template ref) the toolbar formats, undoes and redoes through the sheet's methods (above): the sheet sets the selection and shows a refusal itself. Without it the toolbar calls the source and emits `select` with where the selection goes after an undo or a redo (for the sheet's `focus`) and `error` with a source's refusal. It emits `done` after each call, for the host to give the grid its focus back. Its buttons do not take the focus from the grid. Under 600 px it keeps undo, redo, bold, italic and underline in a row and the rest in a menu; on a narrow frame or under a finger every target is 44 × 44 px.

## The example: an app like Excel

`example/spa.html` is an app like Excel over the local source, all in the browser: New (an empty book in English or Russian), Open (the picker, or an xlsx file dropped on the page), Save as (the book with its values as `<name>.xlsx`, under the name in the bar), a notice of what the file lost, the toolbar (through the sheet's methods), the sheet with its formula bar, tabs, the structure menu on row numbers, column letters and tabs, and selection totals, a status bar (reading the file, computing the book, ready, or calculating while a change is out) and the book's settings (its language, in which formulas, numbers and dates are typed and shown; the book opens again in it, on the sheet it stood on). It speaks Russian or English by the browser's language (`?lang=ru` or `?lang=en` chooses).

A file opens as Excel opens it for its user. xlsx keeps no language, so a file that names none (only xlcalc writes one) opens in the locale of the interface's language, ru-RU under the Russian interface, as Excel shows formulas in its user's language; the settings change it. The app opens the first visible sheet with a value or a formula: Excel opens the sheet active when the file was saved, which xlcalc does not read yet, and a divider sheet at the front would open empty. What the file lost is one line, as Excel's message bar: each kind with its count, what changes values (links to other workbooks, formulas the engine does not parse) first and names last; Details opens each message with its count, the sheets, example cells and names. Reading and writing a file run in a thread of the app (`example/spa-files-worker.ts`), and the book crosses between the threads in slices of 20,000 cells, so a large file does not stop the page while it opens or saves.

`npm run dev:example` serves it at `http://localhost:5181/spa.html`. `npm run build:example` builds it into `example/dist`, which any static server serves (`npx vite preview --config vite.example.config.ts`, or `python3 -m http.server -d example/dist`); there is no server of its own. `npm run check:spa` builds it, serves it and drives it in Chromium at 1920 × 1200 and on a phone at 390 × 844 by touch: a ru-RU book, values and a formula, bold, percent, outside borders, undo, redo, Save as, and the saved file opened again; then the structure menu (a row inserted by a right click and by a long press, a column hidden and shown, a sheet added, renamed and deleted, a column fitted by a double click), with screenshots of each step in `test-results/screens/spa-*.png`.

`npm run check:big` measures the app on large books in Chromium at 1920 × 1200: the production build served by `vite preview` opens each file from its field, and the check measures the first drawn cells of the sheet the app opens and Ready, the jump to the largest sheet, scrolling it down five screens and right three, a number typed into a cell whose dependent is on screen until the dependent shows its new value, Control+Z, Save as (time and size; the saved file is read back and must have the book's sheets and formulas, then it is deleted), the heap of the page and of its Web Workers (CDP), the longest frame gap and long task of each step, and a change of the book's language in its settings. It fails when a file without a language does not open in ru-RU, when the app opens another sheet than the first visible one with data, on a page error or when no edit reaches the screen. Without `BIG_XLSX` it checks itself on xlcalc's synthetic book of 34,000 rows (510,757 cells, written to `test-results`). To run it on your own books, give their paths, joined by `:`; they are read where they lie and nothing is copied:

```sh
BIG_XLSX="$HOME/books/plan.xlsx:$HOME/books/budget.xlsx" npm run check:big
BIG_PROFILE=1 BIG_XLSX=... npm run check:big   # also the page's CPU profiles of each step, test-results/big-*.cpuprofile
```

`npm run check:excel-values` compares the library's values of a book with the values Excel saved in it: each formula cell after a session computes the book against `readXlsx`'s `cached`, numbers within a relative 1e-9. It prints how many differ and exits 1 when any does (`XLSX=a.xlsx:b.xlsx`; a large book wants `NODE_OPTIONS=--max-old-space-size=12288`).

The table goes to the terminal, the numbers to `test-results/big-check.json` and the screenshots to `test-results/screens/big-*.png`. On the synthetic book (Apple M5 Pro, 12.10.2026, the machine busy with other builds) the first sheet is drawn 2.3 to 2.4 s after the file is picked, the dependent shows an edit in 40 to 50 ms and its undo in 25 to 40 ms, Save as takes 4.1 to 4.4 s, and no frame of the page waits longer than 0.3 s.

## A book computed in the browser: `xlcalc-sheet/local`

`localSource(book)` is a `SheetSource` over an xlcalc `BookInput`. It needs `xlcalc` 0.21.1 or later, an optional peer of the package that the host installs. The book computes in a Web Worker (`xlcalc/worker`), or in the page's thread with `calc: 'inline'`. The book's locale module is loaded when the source opens (ru-RU from `xlcalc/locales/ru-RU`); a locale xlcalc has no module for fails the opening with the library's message.

The session does the work as Excel does it, and the source passes it on with every optional method of `SheetSource`. A window is one `getViewport` call: formulas as the book's locale shows them, the text each cell opens with (`editText`), styles, what conditional formats give cells (`appearance`), widths, heights, merged areas and tables. Typed text is read in the book's locale, several cells in one step of the undo (`=СУММ(A1:A2)`, `1 234,5`, `15.03.2026` in ru-RU), and a column a typed number does not fit widens; Delete over a selection clears the content and the hyperlinks, as Excel's Delete. `paste` takes clipboard text by Excel's rules for a paste (an apostrophe stays in the text, a percent cell has no percent editor), `fill` is Control+Return, `undo` and `redo` are the session's, whole. A value a stop rule of data validation does not take is refused before anything is written, and the history stays as it was; a warning's or an information's value is written, as their Yes and OK keep it (the sheet has no alert dialog yet). Structure operations, widths and heights, formatting over each cell's look (`format`, as Excel's Bold over a selection), `explain`, totals, search, the display filter and sort complete it. `openXlsx` reads a file into a book and `saveXlsx` writes the source's book with its values; with `locale` (the language the person chose: a tag, its module loaded on demand, or a locale object) their issues' messages are in that language, in English without it; a tag xlcalc has no module for fails with the library's message.

```ts
import { SheetView } from 'xlcalc-sheet'
import { localSource, openXlsx, saveXlsx } from 'xlcalc-sheet/local'

const { book, issues } = await openXlsx(file, { locale: 'ru-RU' }) // a File from an input or a drop; issues: what the book cannot hold
const source = localSource(book)                                    // <SheetView :source="source" :locale="source.locale" />
const { data } = await saveXlsx(source)                             // the xlsx bytes, formulas with their values
```

The pool starts its thread as `new Worker(new URL('./thread.js', import.meta.url))`, which Vite and other bundlers follow; where a bundler serves the thread elsewhere, pass `pool: { spawn }`. Besides `SheetSource` the source has `ready`, `locale`, `history` (the steps `undo` and `redo` can take, for a toolbar), `book()`, `snapshot()` (the book and its values taken at one moment, what `writeXlsx` takes; it replaces `result()` of 0.2), `freeze()` and `close()`. It tells its subscribers of every change it makes: `changed` for edits and widths, `reload` for structure, styles, lines, frozen panes and their undo. An edit, a paste, a fill, an undo or a redo in a book with conditional formats sends `repaint` after its `changed`, with the ranges of the book's rules the session gives: their cells may look otherwise anywhere the rules reach (a colour scale recolours as one number changes), and a cell update does not carry a look. The values the source keeps for search and the filter stay then, unless a rule shows numbers by a format of its own, and a display filter keeps its rows. `npm run check:local` opens the example's `local.html` (the synthetic book of 3,400 rows) in Chromium, measures it and takes screenshots; the page also opens an xlsx file from its field and saves the book under the file's name.
