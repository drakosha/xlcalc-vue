# Browser tests

The checks build the example app ([example-app.md](example-app.md)) or the local page, drive them in Chromium through Playwright, and save screenshots under `test-results/screens/`. `npm run screens` takes screenshots of the example with every kind of formatting; along the way it checks that the formula bar's layer of colored references lines up with the input's text to the pixel (single-line and expanded, idle and with a long draft, at both widths) and drives the formula bar with the mouse on the desktop: Ctrl+click on a reference, dragging over the text, a click while the cell editor is open, and the reference underlined with a hand pointer while Cmd or Ctrl is held over it (in either order); on the phone, a long press on a reference goes to its range and a tap on it starts editing. It also checks the zoom at 50%, 100% and 200% on the frozen sheet scrolled down and right (`*-70-zoom-*.png`): a click and a tap select the cell under them in the body and in the frozen column, and on the desktop a range dragged in Point mode, the editor over its cell and a column edge dragged by 40 pixels (40 divided by the zoom in the workbook's pixels); Ctrl+wheel's 15% steps on the desktop, a two-finger pinch on the phone (`390-71-pinch.png`), and the Watch Window (`*-72-watch.png`): Add Watch over a range, a formula in its column, a value following another tab's edit, and a double click going to the cell from another sheet.

## App scenarios: `check:spa`

`npm run check:spa` builds the example app, serves it and drives it in Chromium at 1920 × 1200 and on a phone at 390 × 844 with touch input, saving screenshots of each step in `test-results/screens/spa-*.png`:

1. a ru-RU workbook: values and a formula, bold, percentage, outside borders, undo, redo, Save as, and the saved file opened again;
2. the structure menu: a row inserted by a right click and by a long press, a column hidden and unhidden, a sheet added, renamed and deleted, a column fitted by a double click;
3. the fill handle and cut, copy and paste: a formula dragged three rows down by its handle, a range with a formula copied and pasted, a formula cut and pasted, each undone; Ctrl+C, Ctrl+X and Ctrl+V on the desktop, the cell context menu on the phone;
4. Formula AutoComplete: `=СУ` in a cell on the desktop and in the formula bar on the phone, СУММ chosen from the list by Tab or a tap, and the sum checked;
5. Insert Function: fx, СУММ found with its syntax and description, the logical functions by category, OK, `1;2` and the formula bar's ✓ prompting about the unclosed parenthesis, Fix, `)` and ✓ giving 3 with A1 still active, and СУММ first among the recently used;
6. a file with a circular reference opened from the file picker, which keeps its stored values through an edit elsewhere;
7. a file opened on the sheet it saved as active, Save as after a move to another sheet recording that sheet, and a file without a saved active sheet opened on its first visible one.

`DEMO=1 npm run check:spa` runs the whole check on the demo build (`npm run build:demo`) served under `/demo/`, and also checks that `/demo/` opens the app and that the build contains no source maps.

## Large workbooks: `check:big`

`npm run check:big` measures the app on large workbooks in Chromium at 1920 × 1200. The production build served by `vite preview` opens each file from its file input, and the check measures the first drawn cells of the sheet the app opens and Ready; the jump to the largest sheet; scrolling it down five screens and right three; a number typed into a cell whose dependent is on screen, until the dependent displays its new value; Ctrl+Z; Save as (time and size; the saved file is read back, must contain the workbook's sheets and formulas, and is then deleted); the heap of the page and of its Web Workers (CDP); the longest frame gap and long task of each step; and a change of the workbook's language in its settings. It fails when a file without a language does not open in ru-RU, when the app opens a sheet other than the one Excel opens (the saved active sheet, otherwise the first visible one), on a page error, or when no edit reaches the screen. Without `BIG_XLSX` it runs on xlcalc's synthetic workbook of 34,000 rows (510,757 cells, written to `test-results`). To run it on your own workbooks, pass their paths separated by `:`; they are read in place and nothing is copied:

```sh
BIG_XLSX="$HOME/books/plan.xlsx:$HOME/books/budget.xlsx" npm run check:big
BIG_PROFILE=1 BIG_XLSX=... npm run check:big   # also the page's CPU profiles of each step, test-results/big-*.cpuprofile
BIG_URL=https://xlcalc.pridelab.dev/demo/ npm run check:big   # a published app instead of the build, the synthetic workbook only
SPA_URL=https://xlcalc.pridelab.dev/demo/ npm run check:spa   # the app's scenarios on a published app
```

The table is printed to the terminal, the numbers go to `test-results/big-check.json` and the screenshots to `test-results/screens/big-*.png`. On the synthetic workbook (Apple M5 Pro, 12.10.2026, with the machine busy with other builds) the first sheet is drawn 2.3 to 2.4 s after the file is chosen, the dependent displays an edit in 40 to 50 ms and its undo in 25 to 40 ms, Save as takes 4.1 to 4.4 s, and no frame of the page waits longer than 0.3 s.

## Values against Excel: `check:excel-values` and `check:compare-values`

`npm run check:excel-values` compares the library's values for a workbook with the values Excel saved in it: each formula cell, after a session opened with the file's stored values has calculated the workbook (as the app opens a file, so a circular reference keeps them), against `readXlsx`'s `cached`, with numbers compared within a relative 1e-9. It prints how many cells differ and exits with 1 if any do (`XLSX=a.xlsx:b.xlsx`; a large workbook needs `NODE_OPTIONS=--max-old-space-size=12288`).

`npm run check:compare-values` compares the values the app saved in a file with those Excel saved in its copy after a full recalculation (`SAVED=app.xlsx EXCEL=excel.xlsx`): every cell with a value in either file, where a value present in only one file is a difference, excluding cells with volatile formulas; numbers within 1e-12 absolute or 1e-14 relative, dates as serial numbers, text exactly, errors by code. It exits with 1 if any cell differs.

## The local page: `check:local`

`npm run check:local` opens the example's `local.html` (the synthetic workbook of 3,400 rows) in Chromium, measures it and takes screenshots; the page also opens an xlsx file from its file input and saves the workbook under the file's name.
