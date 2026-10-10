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

The runtime depends only on `vue`; colours and fonts are `--xs-*` CSS variables. `npm run dev:example` serves the example with every kind of formatting, `npm run screens` takes its screenshots, `docs/host-parity.md` maps the component onto the the app's grid.

## A book computed in the browser: `xlcalc-sheet/local`

`localSource(book)` is a `SheetSource` over an xlcalc `BookInput`: the book computes in a Web Worker (`xlcalc/worker`), or in the page's thread with `calc: 'inline'`, and the sheet gets every optional method: typed edits read by `parseEntry` in the book's locale (`=СУММ(A1:A2)`, `1,5`, `15.03.2026` in ru-RU), structure, column widths and row heights, formatting, `explain`, totals, search, the display filter and sort. `openXlsx` reads a file into a book and `saveXlsx` writes the source's book with its values. This entry needs `xlcalc` installed (an optional peer of the package).

```ts
import { SheetView } from 'xlcalc-sheet'
import { localSource, openXlsx, saveXlsx } from 'xlcalc-sheet/local'

const { book, issues } = await openXlsx(file) // a File from an input or a drop; issues: what the book cannot hold
const source = localSource(book)               // <SheetView :source="source" :locale="source.locale" />
const { data } = await saveXlsx(source)        // the xlsx bytes, formulas with their values
```

The pool starts its thread as `new Worker(new URL('./thread.js', import.meta.url))`, which Vite and other bundlers follow; where a bundler serves the thread elsewhere, pass `pool: { spawn }`. Besides `SheetSource` the source has `ready`, `locale`, `book()`, `result()`, `freeze()` and `close()`, and tells its subscribers of every change it makes: `changed` for edits and widths, `reload` for structure, styles and frozen panes. `npm run check:local` opens the example's `local.html` (the synthetic book of 3,400 rows) in Chromium, measures it and takes screenshots.
