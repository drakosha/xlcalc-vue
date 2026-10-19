# Example app

`example/spa.html` is an Excel-like app over the local source, running entirely in the browser. It has:

- New (an empty workbook in English or Russian), Open (the file picker, or an xlsx file dropped on the page) and Save as (the workbook with its values as `<name>.xlsx`, under the name in the top bar);
- a notice of what the file lost;
- the toolbar (through the sheet's methods) with Excel's Format Cells dialog;
- the sheet with its formula bar, tabs, the structure menu on row numbers, column letters and tabs, the cell context menu, the fill handle and selection totals;
- a status bar (reading the file, calculating the workbook, ready, or calculating while a change is in progress);
- the workbook's settings: its language, in which formulas, numbers and dates are typed and displayed; the workbook reopens in it, on the sheet that was active.

The interface is in Russian or English according to the browser's language (`?lang=ru` or `?lang=en` chooses one). A single change can write at most 100,000 cells with a value or a formula (xlcalc's `limits.writes`): a cell copied over an entire column, which is 1,048,576 cells and would take seconds and hundreds of megabytes in the worker thread, is rejected immediately with the library's message. Excel has no such limit.

A file opens as Excel opens it for its user. xlsx stores no language, so a file that specifies none (only xlcalc writes one) opens in the locale of the interface language, ru-RU with the Russian interface, as Excel displays formulas in its user's language; the settings change it. The app opens the sheet that was active when the file was saved, as Excel does, otherwise the first visible one (Excel writes no active sheet when the first one was active), and Save as records the sheet on screen as the active one. What the file lost is shown in one line, like Excel's message bar: each kind of issue with its count, those that change values (links to other workbooks, formulas the engine cannot parse) first and names last; Details expands each message with its count, the sheets, example cells and names. Reading and writing a file run in a worker thread of the app (`example/spa-files-worker.ts`), and the workbook and the values stored in the file are transferred between the threads in slices of 20,000 cells, so a large file does not block the page while it opens or saves; the workbook opens with those values, so a circular reference keeps the values Excel displayed.

It needs no server of its own: the dev server serves it, and the build goes to `example/dist`, which any static server can serve.

```sh
npm run dev:example                                  # http://localhost:5181/spa.html
npm run build:example                                # example/dist
npx vite preview --config vite.example.config.ts     # or: python3 -m http.server -d example/dist
```

`npm run build:demo` builds the app alone, for `xlcalc.pridelab.dev/demo/`, into `example/dist-demo`: with its paths under `/demo/` (the worker threads too, which load via `import.meta.url`), with the app itself as `index.html`, so `/demo/` opens it, and without source maps. Any static server that serves the directory at `/demo/` can host it.

The checks that drive the app in a browser are in [checks.md](checks.md).
