# Example app

`example/spa.html` is an app like Excel over the local source, all in the browser: New (an empty book in English or Russian), Open (the picker, or an xlsx file dropped on the page), Save as (the book with its values as `<name>.xlsx`, under the name in the bar), a notice of what the file lost, the toolbar (through the sheet's methods) with Excel's Format Cells dialog, the sheet with its formula bar, tabs, the structure menu on row numbers, column letters and tabs, the cells' menu, the fill handle and selection totals, a status bar (reading the file, computing the book, ready, or calculating while a change is out) and the book's settings (its language, in which formulas, numbers and dates are typed and shown; the book opens again in it, on the sheet it stood on). It speaks Russian or English by the browser's language (`?lang=ru` or `?lang=en` chooses). One change may write at most 100,000 cells with a value or a formula (xlcalc's `limits.writes`): a cell copied over a whole column, 1,048,576 cells and seconds and hundreds of megabytes in the thread, is refused at once with the library's message. Excel has no such limit.

A file opens as Excel opens it for its user. xlsx keeps no language, so a file that names none (only xlcalc writes one) opens in the locale of the interface's language, ru-RU under the Russian interface, as Excel shows formulas in its user's language; the settings change it. The app opens the first visible sheet with a value or a formula: Excel opens the sheet active when the file was saved, which xlcalc does not read yet, and a divider sheet at the front would open empty. What the file lost is one line, as Excel's message bar: each kind with its count, what changes values (links to other workbooks, formulas the engine does not parse) first and names last; Details opens each message with its count, the sheets, example cells and names. Reading and writing a file run in a thread of the app (`example/spa-files-worker.ts`), and the book and the values the file stores cross between the threads in slices of 20,000 cells, so a large file does not stop the page while it opens or saves; the book opens with those values, so a circular reference holds what Excel showed.

It has no server of its own: the dev server serves it, and the build goes to `example/dist`, which any static server serves.

```sh
npm run dev:example                                  # http://localhost:5181/spa.html
npm run build:example                                # example/dist
npx vite preview --config vite.example.config.ts     # or: python3 -m http.server -d example/dist
```

`npm run build:demo` builds the app alone for `xlcalc.pridelab.dev/demo/` into `example/dist-demo`: its paths under `/demo/` (the threads too, which load by `import.meta.url`), its `index.html` the app itself, so `/demo/` opens it, and no source maps. Any static server that serves the directory at `/demo/` serves it.

The checks that drive the app in a browser are in [checks.md](checks.md).
