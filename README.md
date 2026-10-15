# xlcalc-sheet

A Vue 3 sheet component for [xlcalc](https://github.com/drakosha/xlcalc) workbooks: Excel in the browser, with the book's styles, formulas, editing, menus, Find and Replace, data validation and the outline. Where the sheet and Excel could differ, it does what xlcalc measured Excel doing.

The component talks to its data only through a `SheetSource`. Two sources come with it: `memorySource`, a book held in memory for examples and tests, and `localSource` from `xlcalc-sheet/local`, an xlcalc book computed in a Web Worker, with xlsx in and out. A source on a server implements the same interface. MIT.

## Install

```sh
npm install xlcalc-sheet vue
npm install xlcalc        # only for xlcalc-sheet/local, 0.40.0 or later
```

The main entry depends at runtime on `vue` alone and carries no engine.

## Quick start

```vue
<script setup lang="ts">
import { SheetView, memorySource } from 'xlcalc-sheet'
import 'xlcalc-sheet/style.css'

const source = memorySource({ sheets: [{ name: 'Sheet1', cells: { A1: 'Hello', B1: 42 } }] })
</script>

<template>
  <SheetView :source="source" style="height: 480px" />
</template>
```

A book from an xlsx file, computed in the browser:

```ts
import { localSource, openXlsx, saveXlsx } from 'xlcalc-sheet/local'

const { book, cached } = await openXlsx(file, { locale: 'ru-RU' }) // a File from an input or a drop
const source = localSource(book, { session: { cached } })           // <SheetView :source="source" :locale="source.locale" />
const { data } = await saveXlsx(source)                             // the xlsx bytes, formulas with their values
```

Excel's formatting toolbar (`SheetToolbar`) and Format Cells dialog (`FormatDialog`) go over the same source: see [toolbar.md](docs/toolbar.md).

Every feature beyond reading windows is an optional method of the source: without `edit` the sheet is read-only, without `find` there is no search, and so on. The interface speaks English by default and Russian with `ruMessages`. Colours and sizes are `--xs-*` CSS variables.

## Documentation

- [API reference](docs/api.md): components, the sheet's methods, the `SheetSource` contract, `memorySource`, `xlcalc-sheet/local`, CSS variables.
- [Editing, keys and the source](docs/editing.md): Excel's keys, Point mode, what each optional method of the source turns on.
- [Rows, columns and sheets](docs/sheets.md): the menus of lines and tabs, sizes, the outline.
- [Copy, cut and paste; the fill handle](docs/clipboard.md)
- [Reapply, Freeze Panes and the precedents](docs/view.md)
- [Find and Replace](docs/find-replace.md)
- [Data validation](docs/validation.md)
- [Formula AutoComplete](docs/autocomplete.md)
- [The formatting toolbar and Format Cells](docs/toolbar.md)
- [A book computed in the browser](docs/local.md): what `xlcalc-sheet/local` does and how, xlsx files.
- [List mode](docs/list-mode.md): a host's list (a view, a report) on the sheet, sorted by the host, with drill-down and totals.
- [The example app](docs/example-app.md) and [checks in a browser](docs/checks.md)

## Development

```sh
npm install
npm test               # unit tests
npm run typecheck
npm run build          # dist/
npm run dev:example    # the example app at http://localhost:5181/spa.html
```

## License

MIT
