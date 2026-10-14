# A book computed in the browser: `xlcalc-sheet/local`

`localSource(book)` is a `SheetSource` over an xlcalc `BookInput`. It needs `xlcalc` 0.37.0 or later, an optional peer of the package that the host installs. The book computes in a Web Worker (`xlcalc/worker`), or in the page's thread with `calc: 'inline'`. The book's locale module is loaded when the source opens (ru-RU from `xlcalc/locales/ru-RU`); a locale xlcalc has no module for fails the opening with the library's message.

## What the source does

The session does the work as Excel does it, and the source passes it on with every optional method of `SheetSource`. A window is one `getViewport` call: formulas as the book's locale shows them, the text each cell opens with (`editText`), styles, what conditional formats give cells (`appearance`), widths, heights, merged areas and tables.

## Editing

Typed text is read in the book's locale, several cells in one step of the undo (`=СУММ(A1:A2)`, `1 234,5`, `15.03.2026` in ru-RU), and a column a typed number does not fit widens; a number typed over a percent cell takes Excel's `%` (6 is 6%), the cell's own text edited in place (`inPlace`: F2, a double click, the formula bar) does not (6% is 0.06); Delete over a selection clears the content and the hyperlinks, as Excel's Delete. `paste` takes clipboard text by Excel's rules for a paste (an apostrophe stays in the text, a percent cell has no percent editor), `fill` is Control+Return, `undo` and `redo` are the session's, whole.

## Data validation

`validation` gives the session's rule over a cell (`getValidation`): a list's items as the cells show them, by their number format, when the rule shows its dropdown, and the input message unless the rule turns it off. `pick` writes an item with the session's `setCells` as Excel writes it (xlcalc 0.29 measured it in Excel for Mac): a list of cells gives the cell's value as it is, a literal list's item is read as typed in en-US in any book (`"001,1/2,x"` gives 1 and the 2nd of January in `d-mmm`); the list is read again first, and an item whose text is no longer at its place is refused (`conflict`). A value a rule of data validation of any style does not take is refused with the rule's alert before anything is written, and the history stays as it was; several cells at once are refused by the strictest rule they break, as xlcalc's `refuseInvalid` picks it, and Control+Return by the active cell's rule. With `accept` the texts are written without the check, as Excel writes them after Yes or OK; a rule with `showError: false` takes any value, as Excel's does.

## Merging, structure and the rest

Merge Cells and Unmerge Cells are xlcalc's `mergeCells` and `unmergeCells` structure operations (xlcalc 0.37): the area keeps its top-left value, its cells take the top-left cell's look, and a range that meets a table or an array formula of the old kind is refused with the library's message. Structure operations, widths and heights, formatting over each cell's look (`format`, as Excel's Bold over a selection), frozen panes (`freeze`, one `setLayout` call and one step of the undo), `explain`, `end` (the session's `getEnd`, Excel's Range.End as xlcalc measured it), totals, Find and Replace, the display filter and sort complete it.

## Find and Replace

Find is the session's `findCells` with the options of Excel's dialog, which gives the cells Excel's Find gives, in its order, as xlcalc measured Find in Excel for Windows: `?` stands for one UTF-16 unit (`a??b` finds `a😀b`), Look in formulas reads the text a cell opens with but a text without its apostrophe (`'007` is found by `007`, not by `'007`), without Match case `σ`, `ς` and `Σ` are one letter and `i` is not `İ`, and hidden rows and columns are skipped in values always and in formulas on a sheet with a filter (xlcalc's `SheetInput.filtered`), where the hits come by the blocks the hidden lines leave. The session keeps the texts it searches and brings them up to date with the edits: on the data sheet of the synthetic book of 34,000 rows in Node the first search takes some 0.48 s, the same query again or a longer one as it is typed 20 to 35 ms. Under the display filter its rows go to the library (`rows`), a few milliseconds; `after` is Find After, the next hits from a cell. A hit's `text` is what Find read: the shown text, or in formulas the text the cell opens with. Replace All is the session's `replaceInCells`, Excel's Replace All as xlcalc measured it, in one step of undo: the cells Find finds in formulas, whatever Look in says, each with the matches replaced (`*` takes as little as it can, the last one as much), entered again even when the text stays, without the percent editor (5 in a percent cell is 500%); a cell whose new formula would not read is left as it was and comes back in `refused` with why, the others are written.

## Outline

A sheet's `outline` is xlcalc's `outlineGroups`: a group of level L is a longest run of lines at L or deeper, collapsed when Excel keeps the mark on its summary line, which it sets as soon as the lines under it are all hidden, by hand too, and keeps on a summary line hidden inside a hidden group. `outline(change)` is one of xlcalc's structure operations (`editStructure`), one step of the undo with SUBTOTAL and AGGREGATE recomputed, by the rules xlcalc measured in Excel for Windows (its `docs/outline.md`): Group puts each line a level deeper, those at the seventh staying there, and is refused when all of them are; Ungroup takes a level off the lines that have one and leaves hidden lines hidden; a group's − and + are Hide and Show Detail at its summary line, which act on that line's whole detail, so the button of an inner group that shares its summary line with the outer one hides and shows both; Show Detail shows the lines hidden by hand too, all but those of the groups inside whose summary line keeps its mark; a level button hides the lines of that level and deeper, shows the other grouped lines, those hidden by hand among them, and leaves lines outside groups as they are. A refusal is the library's, in the book's language. A group at the sheet's edge (rows 1:3 with summaries above) has no summary line and, as in Excel, no button: the grid draws none, and a command for it is refused. `memorySource` keeps the component's earlier rules (`src/memory-outline.ts`), which miss the shared summary line and the marks set by hand and kept under a level button.

## Files

`openXlsx` reads a file into a book and the values it stores (`cached`), and `saveXlsx` writes the source's book with its values; with `locale` (the language the person chose: a tag, its module loaded on demand, or a locale object) their issues' messages are in that language, in English without it; a tag xlcalc has no module for fails with the library's message.

```ts
import { SheetView } from 'xlcalc-sheet'
import { localSource, openXlsx, saveXlsx } from 'xlcalc-sheet/local'

const { book, cached, issues } = await openXlsx(file, { locale: 'ru-RU' }) // a File from an input or a drop; issues: what the book cannot hold
const source = localSource(book, { session: { cached } })                   // <SheetView :source="source" :locale="source.locale" />
const { data } = await saveXlsx(source)                                     // the xlsx bytes, formulas with their values
```

A book read from a file opens as Excel opens the file with the values it stores, given to the session as `session: { cached }` (xlcalc's `SessionOptions.cached`, 0.27 or later): a circular reference without iterative calculation is not computed, and its cells and those that read them hold the values the file stores, through edits elsewhere, until a value typed into the cycle breaks it (xlcalc's `docs/cycles.md`). Without them those cells hold 0. The session's book keeps what they hold (`held`), so a book taken with `book()` opens holding it again, and `saveXlsx` stores it.

## Threads, the book and events

The pool starts its thread as `new Worker(new URL('./thread.js', import.meta.url))`, which Vite and other bundlers follow; where a bundler serves the thread elsewhere, pass `pool: { spawn }`. Besides `SheetSource` the source has `ready`, `locale`, `history` (the steps `undo` and `redo` can take, for a toolbar), `book()`, `snapshot()` (the book and its values taken at one moment, what `writeXlsx` takes; it replaces `result()` of 0.2) and `close()`. It tells its subscribers of every change it makes: `changed` for edits and widths, `reload` for structure, styles, lines, frozen panes and their undo. An edit, a paste, a fill, an undo or a redo in a book with conditional formats sends `repaint` after its `changed`, with the ranges of the book's rules the session gives: their cells may look otherwise anywhere the rules reach (a colour scale recolours as one number changes), and a cell update does not carry a look. The values the source keeps for search and the filter stay then, unless a rule shows numbers by a format of its own, and a display filter keeps its rows.
