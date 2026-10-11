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

`localSource(book)` is a `SheetSource` over an xlcalc `BookInput`. It needs `xlcalc` 0.19 or later, an optional peer of the package that the host installs. The book computes in a Web Worker (`xlcalc/worker`), or in the page's thread with `calc: 'inline'`. The book's locale module is loaded when the source opens (ru-RU from `xlcalc/locales/ru-RU`); a locale xlcalc has no module for fails the opening with the library's message.

The session does the work as Excel does it, and the source passes it on with every optional method of `SheetSource`. A window is one `getViewport` call: formulas as the book's locale shows them, the text each cell opens with (`editText`), styles, widths, heights, merged areas and tables. Typed text is read in the book's locale, several cells in one step of the undo (`=СУММ(A1:A2)`, `1 234,5`, `15.03.2026` in ru-RU), and a column a typed number does not fit widens; Delete over a selection clears the content and the hyperlinks, as Excel's Delete. `paste` takes clipboard text by Excel's rules for a paste (an apostrophe stays in the text, a percent cell has no percent editor), `fill` is Control+Return, `undo` and `redo` are the session's, whole. A value a stop rule of data validation does not take is refused and taken back. Structure operations, widths and heights, formatting over each cell's look (`format`, as Excel's Bold over a selection), `explain`, totals, search, the display filter and sort complete it. `openXlsx` reads a file into a book and `saveXlsx` writes the source's book with its values.

```ts
import { SheetView } from 'xlcalc-sheet'
import { localSource, openXlsx, saveXlsx } from 'xlcalc-sheet/local'

const { book, issues } = await openXlsx(file) // a File from an input or a drop; issues: what the book cannot hold
const source = localSource(book)               // <SheetView :source="source" :locale="source.locale" />
const { data } = await saveXlsx(source)        // the xlsx bytes, formulas with their values
```

The pool starts its thread as `new Worker(new URL('./thread.js', import.meta.url))`, which Vite and other bundlers follow; where a bundler serves the thread elsewhere, pass `pool: { spawn }`. Besides `SheetSource` the source has `ready`, `locale`, `history` (the steps `undo` and `redo` can take, for a toolbar), `book()`, `snapshot()` (the book and its values taken at one moment, what `writeXlsx` takes; it replaces `result()` of 0.2), `freeze()` and `close()`. It tells its subscribers of every change it makes: `changed` for edits and widths, `reload` for structure, styles, lines, frozen panes and their undo. `npm run check:local` opens the example's `local.html` (the synthetic book of 3,400 rows) in Chromium, measures it and takes screenshots; the page also opens an xlsx file from its field and saves the book under the file's name.
