# Checks in a browser

The checks build the example app ([example-app.md](example-app.md)) or the local page, drive them in Chromium through Playwright and put screenshots under `test-results/screens/`. `npm run screens` takes the screenshots of the example with every kind of formatting; on the way it checks that the formula bar's layer of coloured references lies on the field's text to the pixel (one line and expanded, idle and with a long draft, at both widths) and drives the bar by the mouse on the desktop: Ctrl+click on a reference, a drag over the text, a click while the cell's field is open.

## The app's scenarios: `check:spa`

`npm run check:spa` builds the example app, serves it and drives it in Chromium at 1920 × 1200 and on a phone at 390 × 844 by touch: a ru-RU book, values and a formula, bold, percent, outside borders, undo, redo, Save as, and the saved file opened again; then the structure menu (a row inserted by a right click and by a long press, a column hidden and shown, a sheet added, renamed and deleted, a column fitted by a double click); then the fill handle and cut, copy and paste (a formula dragged three rows down by its handle, a range with a formula copied and pasted, a formula cut and pasted, each undone; Control+C, X and V on the desktop, the cells' menu on the phone); then Formula AutoComplete (`=СУ` in a cell on the desktop and in the formula bar on the phone, СУММ taken from the list by Tab or a tap, the sum checked); then a file with a circular reference opened from the picker, which holds the values it stores through an edit elsewhere; with screenshots of each step in `test-results/screens/spa-*.png`.

`DEMO=1 npm run check:spa` takes the whole check on the demo build (`npm run build:demo`), served under `/demo/`, and checks that `/demo/` is the app and that the build holds no source map.

## Large books: `check:big`

`npm run check:big` measures the app on large books in Chromium at 1920 × 1200: the production build served by `vite preview` opens each file from its field, and the check measures the first drawn cells of the sheet the app opens and Ready, the jump to the largest sheet, scrolling it down five screens and right three, a number typed into a cell whose dependent is on screen until the dependent shows its new value, Control+Z, Save as (time and size; the saved file is read back and must have the book's sheets and formulas, then it is deleted), the heap of the page and of its Web Workers (CDP), the longest frame gap and long task of each step, and a change of the book's language in its settings. It fails when a file without a language does not open in ru-RU, when the app opens another sheet than the first visible one with data, on a page error or when no edit reaches the screen. Without `BIG_XLSX` it checks itself on xlcalc's synthetic book of 34,000 rows (510,757 cells, written to `test-results`). To run it on your own books, give their paths, joined by `:`; they are read where they lie and nothing is copied:

```sh
BIG_XLSX="$HOME/books/plan.xlsx:$HOME/books/budget.xlsx" npm run check:big
BIG_PROFILE=1 BIG_XLSX=... npm run check:big   # also the page's CPU profiles of each step, test-results/big-*.cpuprofile
BIG_URL=https://xlcalc.pridelab.dev/demo/ npm run check:big   # a published app instead of the build, the synthetic book only
SPA_URL=https://xlcalc.pridelab.dev/demo/ npm run check:spa   # the app's scenarios on a published app
```

The table goes to the terminal, the numbers to `test-results/big-check.json` and the screenshots to `test-results/screens/big-*.png`. On the synthetic book (Apple M5 Pro, 12.10.2026, the machine busy with other builds) the first sheet is drawn 2.3 to 2.4 s after the file is picked, the dependent shows an edit in 40 to 50 ms and its undo in 25 to 40 ms, Save as takes 4.1 to 4.4 s, and no frame of the page waits longer than 0.3 s.

## Values against Excel: `check:excel-values` and `check:compare-values`

`npm run check:excel-values` compares the library's values of a book with the values Excel saved in it: each formula cell after a session opened with the file's stored values computes the book (as the app opens a file, so a circular reference holds them) against `readXlsx`'s `cached`, numbers within a relative 1e-9. It prints how many differ and exits 1 when any does (`XLSX=a.xlsx:b.xlsx`; a large book wants `NODE_OPTIONS=--max-old-space-size=12288`).

`npm run check:compare-values` compares the values the app saved in a file with those Excel saved in its copy after a full recalculation (`SAVED=app.xlsx EXCEL=excel.xlsx`): every cell with a value in either file, a value in one only differing, cells of volatile formulas left out; numbers within 1e-12 or a relative 1e-14, dates as serial numbers, text exactly, errors by code. It exits 1 when any differs.

## The local page: `check:local`

`npm run check:local` opens the example's `local.html` (the synthetic book of 3,400 rows) in Chromium, measures it and takes screenshots; the page also opens an xlsx file from its field and saves the book under the file's name.
