# @xlcalc/vue

A Vue 3 spreadsheet component for [xlcalc](https://github.com/drakosha/xlcalc) workbooks: Excel in the browser, with the workbook's styles, formulas, editing, context menus, Find and Replace, data validation and outline. Wherever the component could behave differently from Excel, it follows Excel's behavior as measured by xlcalc.

The component accesses its data only through a `SheetSource`. Two sources are included: `memorySource`, an in-memory workbook for examples and tests, and `localSource` from `@xlcalc/vue/local`, which calculates an xlcalc workbook in a Web Worker and reads and writes xlsx files. A server-side source implements the same interface. MIT license.

## Install

```sh
npm install @xlcalc/vue vue
npm install xlcalc        # for @xlcalc/vue/local, 0.46.0 or later
```

At runtime the main entry point depends only on `vue` and does not include the calculation engine. Its types use `xlcalc/contract`, so a TypeScript project that uses only the main entry point installs `xlcalc` as a dev dependency: `npm install -D xlcalc`.

## Quick start

```vue
<script setup lang="ts">
import { SheetView, memorySource } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({ sheets: [{ name: 'Sheet1', cells: { A1: 'Hello', B1: 42 } }] })
</script>

<template>
  <SheetView :source="source" style="height: 480px" />
</template>
```

A workbook from an xlsx file, calculated in the browser:

```ts
import { localSource, openXlsx, saveXlsx } from '@xlcalc/vue/local'

const { book, cached } = await openXlsx(file, { locale: 'ru-RU' }) // a File from a file input or a drop
const source = localSource(book, { session: { cached } })           // <SheetView :source="source" :locale="source.locale" />
const { data } = await saveXlsx(source)                             // the xlsx bytes: formulas along with their values
```

Excel's formatting toolbar (`SheetToolbar`) and Format Cells dialog (`FormatDialog`) work with the same source; see [toolbar.md](docs/toolbar.md).

Apart from the two required methods that read sheets and cells, every feature is an optional method of the source: without `edit` the sheet is read-only, without `find` there is no search, and so on. The interface is in English by default and in Russian with `ruMessages`. Colors and sizes are set by `--xs-*` CSS variables.

More of these, from a table the page already has to a workbook on your server, are in the [quick start](docs/quick-start.md).

## Documentation

- [Quick start](docs/quick-start.md): common uses, each with a component that works as written.
- [API reference](docs/api.md): components, the sheet's methods, the `SheetSource` contract, `memorySource`, `@xlcalc/vue/local`, CSS variables.
- [Cell editing and keyboard shortcuts](docs/editing.md): Excel's keyboard shortcuts, Point mode, and the features each optional source method enables.
- [Rows, columns and sheets](docs/sheets.md): row, column and sheet tab context menus, sizes, outline.
- [Clipboard and fill handle](docs/clipboard.md)
- [Freeze Panes, Reapply, Show Formulas and Trace Precedents](docs/view.md)
- [Find and Replace](docs/find-replace.md)
- [Go To and Symbol](docs/go-to-and-symbol.md)
- [Data validation](docs/validation.md)
- [Formula AutoComplete](docs/autocomplete.md)
- [The formatting toolbar and Format Cells](docs/toolbar.md)
- [In-browser calculation](docs/local.md): how `@xlcalc/vue/local` works, xlsx files.
- [List mode](docs/list-mode.md): the host's list (a view or a report) displayed on the sheet, sorted by the host, with drill-down and a totals row.
- [Example app](docs/example-app.md) and [browser tests](docs/checks.md)

## Development

```sh
npm install
npm test               # unit tests
npm run typecheck
npm run build          # dist/
npm run check:package  # publint and attw over the packed package
npm run dev:example    # the example app at http://localhost:5181/spa.html
```

## License

MIT
