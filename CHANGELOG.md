# Changelog

## Unreleased

The first public release. The versions before it were internal builds and were not published to npm; their changes are folded into this entry, and the README with `docs/` describes the package as it is.

What is in it:

- `SheetView`: a windowed grid over a `SheetSource`, with the book's styles, number formats, conditional formats, merged areas, tables, hyperlinks and notes; editing with Excel's keys, the formula bar (one line, expanded by its arrow or Control+Shift+U), Point mode by the mouse and the keys as Excel measured it (xlcalc's `refAt`: an arrow points, enters the formula or does nothing; Command or Control and an arrow, Home and Page Down move the reference as on the sheet), Formula AutoComplete, fill and paste, Control+Return, undo and redo; selection, Control and an arrow as Excel's Range.End, frozen panes, the fill handle going on with series as Excel's AutoFill (`fillRange`: 1, 2 → 3, 4, dates, Item1 → Item2; Ctrl or Option at the release for a series of one number), cut, copy and paste as Excel copies text, Find and Replace, data validation with Excel's dropdown, input message and alerts, the outline, the display filter and sort, the precedents' panel, menus of cells, lines and tabs, column widths and row heights by the mouse with AutoFit on a double click.
- List mode: a host's list on the sheet, its titles sorted by the host (`sort`), drill-down cells (`cellDrill`), rows opened from the first column (`openRows`, `rowOpen`) and a totals row held at the bottom.
- `SheetToolbar` and `FormatDialog`: Excel's formatting row, with Merge & Center and its list (Merge Across, Merge Cells, Unmerge Cells) and Excel's warning before values are lost, and its Format Cells dialog.
- The `SheetSource` contract (`src/source.ts`): every feature beyond reading windows is an optional method, so a source on a server and one in the page give what they have. `functionSignature(name)` gives a function's arguments for Excel's tip while a formula is typed.
- `memorySource`: a source over a book held in memory, for an example or a test.
- `xlcalc-sheet/local`: `localSource`, a source over an xlcalc book that computes in a Web Worker or in the page's thread and does what the grid asks as Excel does it, with `openXlsx` and `saveXlsx`; it needs `xlcalc` 0.39.0 or later, an optional peer.
- `docs/`: the API reference (`docs/api.md`) and a page for each feature.
- The interface's texts in English by default and in Russian (`ruMessages`).
- `example/spa.html`: an app like Excel over `./local`, all in the browser, without a server.
- The main entry depends at runtime on `vue` alone and carries no engine. Its test holds it under 380,000 characters (357,000 now): the engine would add hundreds of thousands.
