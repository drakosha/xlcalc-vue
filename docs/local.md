# In-browser calculation: `@xlcalc/vue/local`

`localSource(book)` is a `SheetSource` over an xlcalc `BookInput`. It requires `xlcalc` 0.46.0 or later, an optional peer dependency of the package that the host installs. Some of its methods need a later xlcalc: `autoSum` and the workbook's default font need 0.59.0 (with an older one there is no AutoSum and the font is Calibri 11), and with 0.59.0 `setNames` refuses a name Excel rejects (an older xlcalc accepts it); Paste Special and Format Painter (`copyRange` with `paste` and `transpose`) need 0.60.0, and an older xlcalc refuses them with its message. The workbook is calculated in a Web Worker (`xlcalc/worker`), or on the page's main thread with `calc: 'inline'`. The workbook's locale module is loaded when the source opens (ru-RU from `xlcalc/locales/ru-RU`); for a locale xlcalc has no module for, opening fails with the library's message.

## What the source does

The xlcalc session does the work as Excel does, and the source exposes it through every optional method of `SheetSource`. A window is one `getViewport` call: formulas as displayed in the workbook's locale, the text each cell opens with for editing (`editText`), styles, the cells' conditional formatting (`appearance`), widths, heights, merged ranges and tables.

## Editing

Typed text is parsed in the workbook's locale, with several cells as a single undo step (`=СУММ(A1:A2)`, `1 234,5`, `15.03.2026` in ru-RU), and a column widens when a typed number does not fit. A number typed over a percentage cell gets Excel's `%` (6 becomes 6%), while the cell's own text edited in place (`inPlace`: F2, a double click, the formula bar) does not (6% stays 0.06). Delete over a selection clears the contents and the hyperlinks, as Excel's Delete does. An unclosed formula is closed or rejected as Excel handles it on Return (xlcalc 0.44): one that starts with a function whose parenthesis is the only one open is closed silently (`=СУММ(1;2`), and any other is `invalid`, with the session's `correction` when Excel suggests one (`=(1+2` becomes `=(1+2)`, `=1+СУММ(2` becomes `=1+СУММ(2)`), in `edit` and `fill` alike. `paste` parses clipboard text by Excel's paste rules (a leading apostrophe stays in the text, and a percentage cell has no automatic percentage entry), `fill` is Ctrl+Enter, and `undo` and `redo` are the session's, undoing whole changes.

## Data validation

`validation` returns the session's rule for a cell (`getValidation`): a list's items as the cells display them, by their number format, when the rule shows its dropdown, and the input message unless the rule turns it off. `pick` writes an item with the session's `setCells` as Excel writes it (measured by xlcalc 0.29 in Excel for Mac): an item from a cell range gives that cell's value as is, and an item of a literal list is parsed as typed in en-US in any workbook (`"001,1/2,x"` gives 1 and January 2 in `d-mmm`); the list is read again first, and an item whose text is no longer at its position is rejected (`conflict`). A value that violates a data validation rule of any style is rejected with the rule's alert before anything is written, and the undo history stays unchanged; several cells at once are rejected by the strictest rule they violate, as xlcalc's `refuseInvalid` chooses it, and Ctrl+Enter by the active cell's rule. With `accept` the texts are written without the check, as Excel writes them after Yes or OK; a rule with `showError: false` accepts any value, as Excel's does.

## Fill, merging, structure and the rest

The fill handle's `fillRange` is xlcalc's AutoFill (xlcalc 0.39, measured in Excel for Windows): the target range is filled from the source range as `copyRange` copies it, then series continue (numbers by their trend, dates by day, by step or by month, times, text with a number, quarters, day and month names), backwards when dragging up or left; the copies also go into hidden rows, on a filtered sheet as well, as Excel's fill handle fills them (from xlcalc 0.51.0; with an earlier xlcalc the rows hidden by a filter stay empty). Merge Cells and Unmerge Cells are xlcalc's `mergeCells` and `unmergeCells` structure operations (xlcalc 0.37): the merged range keeps its top-left value, its cells take the top-left cell's formatting, and a range that overlaps a table or a legacy array formula is rejected with the library's message. Structure operations, widths and heights, formatting on top of each cell's formatting (`format`, as Excel's Bold over a selection), frozen panes (`freeze`, one `setLayout` call and a single undo step), `explain`, `end` (the session's `getEnd`, Excel's Range.End as measured by xlcalc), totals, Find and Replace, the display filter and sort complete the source. Formula AutoComplete's `functionNames` returns every function xlcalc calculates, by its name in the workbook's locale, and the argument tooltip's `functionSignature` is xlcalc's `functionSignature` (xlcalc 0.40) with its `functionDescription` (xlcalc 0.43) in the workbook's locale. The localized names (`xlcalc/function-text/ru-RU` for ru-RU) and the descriptions (`xlcalc/function-text`) load with the first tooltip; a name is accepted in English or in the locale, in any letter case, and the result is null for a function xlcalc does not have ([autocomplete.md](autocomplete.md)).

## Find and Replace

Find is the session's `findCells` with the options of Excel's dialog, which returns the cells Excel's Find returns, in the same order, as measured by xlcalc in Excel for Windows: `?` matches one UTF-16 code unit (`a??b` finds `a😀b`); Look in Formulas searches the text a cell opens with for editing, but without a text's apostrophe (`'007` is found by `007`, not by `'007`); without Match case, `σ`, `ς` and `Σ` are the same letter and `i` does not match `İ`; and hidden rows and columns are always skipped when searching values, and when searching formulas on a sheet with a filter (xlcalc's `SheetInput.filtered`), where the matches come in the blocks between the hidden rows or columns. The session keeps the texts it searches and updates them with edits: on the data sheet of the synthetic workbook of 34,000 rows in Node, the first search takes about 0.48 s, and the same query again, or a longer one as it is typed, 20 to 35 ms. Under the display filter its rows are passed to the library (`rows`), which takes a few milliseconds; `after` is Find After, the next matches after a cell. Replace All is the session's `replaceInCells`, Excel's Replace All as measured by xlcalc, as a single undo step: it takes the cells Find finds in formulas, whatever the Look in setting, replaces the matches in each (each `*` matches as little as possible, the last one as much as possible), and enters each cell again even if its text is unchanged, without automatic percentage entry (5 in a percentage cell becomes 500%); a cell whose new formula would not parse is left unchanged and is returned in `refused` with the reason, and the other cells are written.

## Display filter and totals

The display filter, its sort and the filter funnel's value list follow the same rules as in `memorySource` (`src/sheet-query.ts`): a condition compares the displayed text, the header rows stay on top, and a sort and the funnel's list order values as Excel's SORT does (xlcalc's `valueComparator`): numbers, text case-insensitively in the order of the workbook's locale, FALSE, TRUE, errors, with empty cells and formulas returning `""` last in either direction. `./local` passes the workbook's locale to the comparator and `memorySource` passes none, which means en-US; the ru-RU order is the same as en-US (`Item10` before `Item2`, Latin before Cyrillic), as Excel for Windows sorts in every locale, whereas Excel for Mac in ru-RU puts Cyrillic first. The source reads from the session only the columns a condition, a sort or a funnel needs, the first time they are needed, keeps them in arrays by row and updates them with edits: on the data sheet of the synthetic workbook of 34,000 rows by 15 columns in Node, through the pool, the first sort takes about 0.16 s and 7 MB, and the first funnel 0.12 s. A request's signal stops the session's read of its columns (`getCells` with `signal`), and a read stopped for one request is repeated for another request still waiting for it. Selection totals come from the session's `summarizeRanges`, which counts without transferring the cells, and under a display filter covers only its rows. A cell is counted when it displays something or contains a value, so a number hidden by its format is counted and a formula returning `""` is not, as the grid counts the cells of its windows.

## Outline

A sheet's `outline` comes from xlcalc's `outlineGroups`: a group of level L is a maximal run of rows or columns at level L or deeper, collapsed when Excel keeps the collapsed flag on its summary row or column. Excel sets the flag as soon as all the rows or columns under it are hidden, manually hidden ones included, and keeps it on a summary row or column hidden inside a hidden group. `outline(change)` is one of xlcalc's structure operations (`editStructure`), a single undo step that recalculates SUBTOTAL and AGGREGATE, following the rules xlcalc measured in Excel for Windows (its `docs/outline.md`):

- Group moves each row or column one level deeper, leaving those at the seventh level there, and is rejected when all of them are at the seventh level.
- Ungroup removes a level from the rows or columns that have one and leaves hidden ones hidden.
- A group's − and + buttons are Hide Detail and Show Detail at its summary row or column, which act on all of that summary's detail, so the button of an inner group that shares its summary row with the outer group hides and shows both.
- Show Detail also shows rows or columns hidden manually, except those of inner groups whose summary keeps its collapsed flag.
- A level button hides the rows or columns of that level and deeper, shows the other grouped ones, manually hidden ones included, and leaves rows or columns outside groups as they are.

Rejection messages come from the library, in the workbook's language. A group at the edge of the sheet (rows 1:3 with summaries above) has no summary row and, as in Excel, no button: the grid draws none, and a command for it is rejected. `memorySource` also derives its groups with xlcalc's `outlineGroups` and runs the same commands without the engine, through xlcalc's `outlineEdit` from `xlcalc/a1`, with the same rejections in English; a row or column it hides gets a size of 0, and one it shows again gets the default size.

## Files

`openXlsx` reads a file into a workbook and the values stored in it (`cached`), and `saveXlsx` writes the source's workbook with its values. With `locale` (the language the user chose: a tag whose module is loaded on demand, or a locale object), their issue messages are in that language; without it, in English. A tag xlcalc has no module for fails with the library's message.

A page that opens an xlsx file, reports what the file lost, and saves the workbook back under the file's name:

```vue
<script setup lang="ts">
import { onUnmounted, shallowRef } from 'vue'
import { SheetView } from '@xlcalc/vue'
import { groupIssues, localSource, openXlsx, saveXlsx, type LocalSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = shallowRef<LocalSource | null>(null)
const lost = shallowRef<string[]>([])
let fileName = 'book.xlsx'

async function open(event: Event): Promise<void> {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (!file) return
  // cached: the values stored in the file, which a circular reference keeps; issues: what the workbook cannot keep
  const { book, cached, issues } = await openXlsx(file)
  void source.value?.close()
  source.value = localSource(book, { session: { cached } })
  fileName = file.name
  lost.value = groupIssues(issues).map((group) => `${group.message} (${group.count})`)
}

async function save(): Promise<void> {
  if (!source.value) return
  const { data } = await saveXlsx(source.value) // the xlsx bytes: formulas along with their values
  const link = document.createElement('a')
  link.href = URL.createObjectURL(new Blob([data as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
  link.download = fileName
  link.click()
  setTimeout(() => URL.revokeObjectURL(link.href))
}

onUnmounted(() => void source.value?.close())
</script>

<template>
  <input type="file" accept=".xlsx" @change="open" />
  <button :disabled="!source" @click="save">Save as</button>
  <p v-for="line in lost" :key="line">{{ line }}</p>
  <SheetView v-if="source" :source="source" :locale="source.locale" style="height: 420px" />
</template>
```

A workbook read from a file opens as Excel opens the file, with its stored values passed to the session as `session: { cached }` (xlcalc's `SessionOptions.cached`, 0.27 or later): a circular reference without iterative calculation is not calculated, and its cells and the cells that depend on them keep the values stored in the file, through edits elsewhere, until a value typed into the cycle breaks it (xlcalc's `docs/cycles.md`). Without them those cells contain 0. The session's workbook records the kept values (`held`), so a workbook obtained with `book()` opens with them again, and `saveXlsx` saves them.

## Threads, the workbook and events

The pool starts its thread as `new Worker(new URL('./thread.js', import.meta.url))`, which Vite and other bundlers recognize; where a bundler serves the thread script from elsewhere, pass `pool: { spawn }`. Besides the `SheetSource` methods, the source has `ready`, `locale`, `history` (the steps `undo` and `redo` can take, for a toolbar), `book()`, `snapshot()` (the workbook and its values captured at one moment, which is what `writeXlsx` takes; it replaces `result()` of 0.2) and `close()`. It notifies its subscribers of every change it makes: `changed` for edits and widths, `reload` for structure, number formats, rows and columns, frozen panes and their undo, and `repaint` with the formatted ranges for a style set by `format` (a style changes no text or size). The values the source caches for the display filter are kept through a style change, a freeze, and an undo or redo of an edit, since these change no cells or list the cells they change. An edit, a paste, a fill, an undo or a redo in a workbook with conditional formats sends `repaint` after its `changed`, with the ranges of the workbook's rules returned by the session: their cells may look different anywhere the rules apply (a color scale changes colors when a single number changes), and a cell update does not carry formatting. The values the source caches for search and the filter are kept then, unless a rule displays numbers with its own number format, and a display filter keeps its rows.
