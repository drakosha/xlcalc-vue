# In-browser calculation: `@xlcalc/vue/local`

`localSource(book)` is a `SheetSource` over an xlcalc `BookInput`. It requires `xlcalc` 0.71.2 or later, an optional peer dependency of the package that the host installs: the package is built over 0.71.2 and calls what it and the versions before it added (calculation options, Go To Special, Create from Selection, Paste Special's operations, sheet protection, the AutoFilter and the range it acts on, the sheet's view, Workbook Statistics, CSV files). The workbook is calculated in a Web Worker (`xlcalc/worker`), or on the page's main thread with `calc: 'inline'`. The workbook's locale module is loaded when the source opens (ru-RU from `xlcalc/locales/ru-RU`); for a locale xlcalc has no module for, opening fails with the library's message.

## What the source does

The xlcalc session does the work as Excel does, and the source exposes it through every optional method of `SheetSource`. A window is one `getViewport` call: formulas as displayed in the workbook's locale, the text each cell opens with for editing (`editText`), styles, the cells' conditional formatting (`appearance`), widths, heights, merged ranges and tables.

## Editing

Typed text is parsed in the workbook's locale, with several cells as a single undo step (`=СУММ(A1:A2)`, `1 234,5`, `15.03.2026` in ru-RU), and a column widens when a typed number does not fit. A number typed over a percentage cell gets Excel's `%` (6 becomes 6%), while the cell's own text edited in place (`inPlace`: F2, a double click, the formula bar) does not (6% stays 0.06). Delete over a selection clears the contents and the hyperlinks, as Excel's Delete does. An unclosed formula is closed or rejected as Excel handles it on Return (xlcalc 0.44): one that starts with a function whose parenthesis is the only one open is closed silently (`=СУММ(1;2`), and any other is `invalid`, with the session's `correction` when Excel suggests one (`=(1+2` becomes `=(1+2)`, `=1+СУММ(2` becomes `=1+СУММ(2)`), in `edit` and `fill` alike. `paste` parses clipboard text by Excel's paste rules (a leading apostrophe stays in the text, and a percentage cell has no automatic percentage entry), `fill` is Ctrl+Enter, and `undo` and `redo` are the session's, undoing whole changes.

## Data validation

`validation` returns the session's rule for a cell (`getValidation`): a list's items as the cells display them, by their number format, when the rule shows its dropdown, and the input message unless the rule turns it off. `pick` writes an item with the session's `setCells` as Excel writes it (measured by xlcalc 0.29 in Excel for Mac): an item from a cell range gives that cell's value as is, and an item of a literal list is parsed as typed in en-US in any workbook (`"001,1/2,x"` gives 1 and January 2 in `d-mmm`); the list is read again first, and an item whose text is no longer at its position is rejected (`conflict`). A value that violates a data validation rule of any style is rejected with the rule's alert before anything is written, and the undo history stays unchanged; several cells at once are rejected by the strictest rule they violate, as xlcalc's `refuseInvalid` chooses it, and Ctrl+Enter by the active cell's rule. With `accept` the texts are written without the check, as Excel writes them after Yes or OK; a rule with `showError: false` accepts any value, as Excel's does.

## Fill, merging, structure and the rest

The fill handle's `fillRange` is xlcalc's AutoFill (xlcalc 0.39, measured in Excel for Windows): the target range is filled from the source range as `copyRange` copies it, then series continue (numbers by their trend, dates by day, by step or by month, times, text with a number, quarters, day and month names), backwards when dragging up or left; the copies also go into hidden rows, on a filtered sheet as well, as Excel's fill handle fills them (from xlcalc 0.51.0; with an earlier xlcalc the rows hidden by a filter stay empty). Merge Cells and Unmerge Cells are xlcalc's `mergeCells` and `unmergeCells` structure operations (xlcalc 0.37): the merged range keeps its top-left value, its cells take the top-left cell's formatting, and a range that overlaps a table or a legacy array formula is rejected with the library's message. Structure operations, widths and heights, formatting on top of each cell's formatting (`format`, as Excel's Bold over a selection), frozen panes (`freeze`, one `setLayout` call and a single undo step), the AutoFilter (`setAutoFilter`, one `setLayout` call; removing it from a filtered sheet also shows the rows of its range below the header and clears `filtered`, as Excel's Filter button does, [autofilter.md](autofilter.md)), `explain`, `end` (the session's `getEnd`, Excel's Range.End as measured by xlcalc), totals, Find and Replace, the display filter and sort complete the source. Formula AutoComplete's `functionNames` returns every function xlcalc calculates, by its name in the workbook's locale, and the argument tooltip's `functionSignature` is xlcalc's `functionSignature` (xlcalc 0.40) with its `functionDescription` (xlcalc 0.43) in the workbook's locale. The localized names (`xlcalc/function-text/ru-RU` for ru-RU) and the descriptions (`xlcalc/function-text`) load with the first tooltip; a name is accepted in English or in the locale, in any letter case, and the result is null for a function xlcalc does not have ([autocomplete.md](autocomplete.md)).

## Protection

xlcalc's session does not enforce a sheet's protection, so the source does: each change to a protected sheet is checked before it reaches the session and refused as Excel for Mac refuses it, with Excel's alert in the workbook's language (`code: 'protected'`), deleting a row or a column with a locked cell included, as Excel for Windows words it; `protect` is one `setLayout` call and a single undo step. What it allows, measured in Excel, is in [protection.md](protection.md#local).

## Find and Replace

Find is the session's `findCells` with the options of Excel's dialog, which returns the cells Excel's Find returns, in the same order, as measured by xlcalc in Excel for Windows: `?` matches one UTF-16 code unit (`a??b` finds `a😀b`); Look in Formulas searches the text a cell opens with for editing, but without a text's apostrophe (`'007` is found by `007`, not by `'007`); without Match case, `σ`, `ς` and `Σ` are the same letter and `i` does not match `İ`; and hidden rows and columns are always skipped when searching values, and when searching formulas on a sheet with a filter (xlcalc's `SheetInput.filtered`), where the matches come in the blocks between the hidden rows or columns. The session keeps the texts it searches and updates them with edits: on the data sheet of the synthetic workbook of 34,000 rows in Node, the first search takes about 0.48 s, and the same query again, or a longer one as it is typed, 20 to 35 ms. Under the display filter its rows are passed to the library (`rows`), which takes a few milliseconds; `after` is Find After, the next matches after a cell. Replace All is the session's `replaceInCells`, Excel's Replace All as measured by xlcalc, as a single undo step: it takes the cells Find finds in formulas, whatever the Look in setting, replaces the matches in each (each `*` matches as little as possible, the last one as much as possible), and enters each cell again even if its text is unchanged, without automatic percentage entry (5 in a percentage cell becomes 500%); a cell whose new formula would not parse is left unchanged and is returned in `refused` with the reason, and the other cells are written.

## Display filter and totals

The display filter, its sort and the filter funnel's value list follow the same rules as in `memorySource` (`src/sheet-query.ts`): a condition compares the displayed text, the header rows stay on top, the rows below an AutoFilter's last row (`SheetFilter.last`) stay below in their order, and a sort and the funnel's list order values as Excel's SORT does (xlcalc's `valueComparator`): numbers, text case-insensitively in the order of the workbook's locale, FALSE, TRUE, errors, with empty cells and formulas returning `""` last in either direction. `./local` passes the workbook's locale to the comparator and `memorySource` passes none, which means en-US; the ru-RU order is the same as en-US (`Item10` before `Item2`, Latin before Cyrillic), as Excel for Windows sorts in every locale, whereas Excel for Mac in ru-RU puts Cyrillic first. The source reads from the session only the columns a condition, a sort or a funnel needs, the first time they are needed, keeps them in arrays by row and updates them with edits: on the data sheet of the synthetic workbook of 34,000 rows by 15 columns in Node, through the pool, the first sort takes about 0.16 s and 7 MB, and the first funnel 0.12 s. A request's signal stops the session's read of its columns (`getCells` with `signal`), and a read stopped for one request is repeated for another request still waiting for it. Selection totals come from the session's `summarizeRanges`, which counts without transferring the cells, and under a display filter covers only its rows. The rows of a filter stay as `visibleRows` returned them until the filter is applied again, as Excel keeps them until Reapply, and the source keeps them for each filter (up to eight a sheet), so viewers that share one source on a server with different filters each keep theirs ([remote.md](remote.md)). A cell is counted when it displays something or contains a value, so a number hidden by its format is counted and a formula returning `""` is not, as the grid counts the cells of its windows.

## Outline

A sheet's `outline` comes from xlcalc's `outlineGroups`: a group of level L is a maximal run of rows or columns at level L or deeper, collapsed when Excel keeps the collapsed flag on its summary row or column. Excel sets the flag as soon as all the rows or columns under it are hidden, manually hidden ones included, and keeps it on a summary row or column hidden inside a hidden group. `outline(change)` is one of xlcalc's structure operations (`editStructure`), a single undo step that recalculates SUBTOTAL and AGGREGATE, following the rules xlcalc measured in Excel for Windows (its `docs/outline.md`):

- Group moves each row or column one level deeper, leaving those at the seventh level there, and is rejected when all of them are at the seventh level.
- Ungroup removes a level from the rows or columns that have one and leaves hidden ones hidden.
- A group's − and + buttons are Hide Detail and Show Detail at its summary row or column, which act on all of that summary's detail, so the button of an inner group that shares its summary row with the outer group hides and shows both.
- Show Detail also shows rows or columns hidden manually, except those of inner groups whose summary keeps its collapsed flag.
- A level button hides the rows or columns of that level and deeper, shows the other grouped ones, manually hidden ones included, and leaves rows or columns outside groups as they are.

Rejection messages come from the library, in the workbook's language. A group at the edge of the sheet (rows 1:3 with summaries above) has no summary row and, as in Excel, no button: the grid draws none, and a command for it is rejected. `memorySource` also derives its groups with xlcalc's `outlineGroups` and runs the same commands without the engine, through xlcalc's `outlineEdit` from `xlcalc/a1`, with the same rejections in English; a row or column it hides gets a size of 0, and one it shows again gets the default size.

## Regular expressions

Match mode 3 of `XLOOKUP` and `XMATCH` finds a regular expression; xlcalc reads it as JavaScript does unless a level of `xlcalc/regex` (0.66.0) is registered. The option `regex` loads one before the workbook is calculated, as a chunk of its own that a host without it never downloads: `pcreSyntax` reads Excel's PCRE2 syntax (possessive quantifiers, atomic groups, `(?i)` anywhere) at JavaScript's speed, for workbooks the host trusts; `pcreEngine` runs xlcalc's PCRE2 engine, which also ends a heavy pattern at Excel's match limit (`#VALUE!`), for patterns the users write. xlcalc keeps the last level registered for every calculation of the page, a pool's threads too, so sources that ask for different levels share the last one. xlcalc's `docs/regex.md` compares the levels.

```ts
const source = localSource(book, { regex: 'pcreEngine' }) // XMATCH("(a+)+$", A1, 3) over a long text is #VALUE! at once, as in Excel
```

## PivotTables

A workbook's PivotTables (xlcalc 0.71.0's `SheetInput.pivotTables`, read from a file by `openXlsx`) calculate as cells without xlcalc's PivotTable engine, and `GETPIVOTDATA` reads them; refreshing them, laying them out again and typing over their labels need the engine, `xlcalc/pivot`. The source loads it as a chunk of its own and registers it with `registerPivot`, xlcalc keeping it for every session of the page from then on, a pool's threads too: before the workbook is calculated when the workbook has PivotTables or pivot caches, and otherwise before the first operation on PivotTables (`structure`'s `addPivotTable` and the others), so a host showing workbooks without them never downloads it. A PivotTable added that way is laid out at once; it keeps its cells empty, its cache marked `stale`, only until the page has loaded the engine, which is when the engine cannot be loaded. A workbook whose engine fails to load still opens, its PivotTables' cells with their values and `GETPIVOTDATA` working; Refresh and the operations on PivotTables then refuse with a message for the user. The source lists each sheet's PivotTables with their areas (`SheetInfo.pivots`), answers what their cells refuse with Excel's alerts, renames a label typed over it, refreshes one through `structure` and all of them through Refresh All, and gates them on a protected sheet by Use PivotTable and PivotChart ([pivot-tables.md](pivot-tables.md)).

## Calculation options

The workbook's calculation option (xlcalc's `BookInput.calculation`, 0.66.0) is `calculation`; `setCalculation` changes it and `calculate` is F9 and Shift+F9. xlcalc calculates every edit whatever the option says, so in Manual the source keeps showing the values of the last calculation, puts them back over the session's in each window, and shows the session's on `calculate`; with an xlcalc before 0.66.0 the sheet shows calculated values in Manual too. What the source holds and where it differs from Excel are in [calculation.md](calculation.md#in-xlcalcvuelocal).

## Files

xlsx files are read and written by `openXlsx` and `saveXlsx`, CSV files by `openCsv` and `saveCsv` ([CSV files](#csv-files)).

`openXlsx` reads a file into a workbook and the values stored in it (`cached`), and `saveXlsx` writes the source's workbook with its values. With `locale` (the language the user chose: a tag whose module is loaded on demand, or a locale object), their issue messages are in that language; without it, in English. A tag xlcalc has no module for fails with the library's message.

A workbook does not hold everything a file may: charts, pictures, pivot tables, the theme, page setup. To save it without losing them, open the file with `keep: true` (xlcalc's `openKept`) and save with `over`, the same file's bytes (xlcalc's `writeKept`): the edited workbook is written over the original, which keeps those parts as the file has them, and without edits gives the original bytes back. A change the file cannot take, one that would leave such a part pointing at the wrong cells (a sheet with a pivot table deleted, rows moved under a chart whose references the library cannot write back), is rejected with xlcalc's `KeptRefusal` (`error.name === 'KeptRefusal'`, the reason in its message and its kind in `code`: `links`, `rules`, `moved`, `slicers` (slicers or timelines added or deleted, xlcalc 0.71.0) and the others of xlcalc's `KeptRefusalCode`), and nothing is written; the host may then save the workbook anew, without `over`, which loses those parts. The source does not list those names. `mode: 'values'` is for a new file only. In Manual, `saveXlsx` calculates first, as Excel recalculates a workbook before saving it, and the file keeps the option ([calculation.md](calculation.md#saving)).

A host that stores the workbook `openXlsx` read with `keep: true` (its `book`, or `book()` of a source over it) to save over the file later must open the file again after moving to xlcalc 0.67.0: from 0.67.0 `writeKept` refuses a workbook whose carriers of the file's parts were laid out by another version of the library or for another file, and so every workbook kept from `openKept` 0.66 or earlier, with `KeptRefusal` `code: 'layout'`, until the original file is opened again with `keep: true`. A workbook without such parts (no carriers) is written as before. The refusal is the same as any other: nothing is written, and saving anew, without `over`, still works.

```ts
import { localSource, openXlsx, saveXlsx } from '@xlcalc/vue/local'

export async function editKept(file: File): Promise<Uint8Array> {
  const { book, cached } = await openXlsx(file, { keep: true })
  const source = localSource(book, { session: { cached } })
  await source.edit([{ sheet: book.sheets[0]!.name, cell: 'A1', text: 'Checked' }])
  try {
    return (await saveXlsx(source, { over: file })).data // the charts and the theme as the file has them
  } catch (e) {
    if ((e as Error).name !== 'KeptRefusal') throw e
    return (await saveXlsx(source)).data // anew: what the workbook holds, without those parts
  } finally {
    await source.close()
  }
}
```

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
  // cached: the values stored in the file, which a circular reference keeps; file: its name, for CELL and INFO; issues: what the workbook cannot keep
  const { book, cached, issues, file: named } = await openXlsx(file)
  void source.value?.close()
  source.value = localSource(book, { session: { cached, file: named } })
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

A workbook read from a file opens as Excel opens the file, with its stored values passed to the session as `session: { cached }` (xlcalc's `SessionOptions.cached`): a circular reference without iterative calculation is not calculated, and its cells and the cells that depend on them keep the values stored in the file, through edits elsewhere, until a value typed into the cycle breaks it (xlcalc's `docs/cycles.md`). Without them those cells contain 0. The session's workbook records the kept values (`held`), so a workbook obtained with `book()` opens with them again, and `saveXlsx` saves them.

A formula over a function the engine cannot calculate keeps the value its file stores too, as Excel shows it until it calculates the cell (xlcalc 0.68.0, its `docs/formulas.md`, Values a file stores): one over a function of connections (`CUBEVALUE`, `WEBSERVICE`, `STOCKHISTORY` and the others) always, and one over an add-in's function, which xlcalc does not know, because the source opens the session with `unknownFunctions: 'held'` unless the host's `session` says otherwise (`'name'` makes it `#NAME?` at once, as Excel without the add-in calculates it). The cells reading it are calculated from that value. Once the cell is edited or a cell it reads changes, it is calculated: `#NAME?` for an unknown function, the answer without the connection for a function of connections, or the host's answer ([below](#calls-outside-the-workbook)). `openXlsx` names a `File` it reads (`file: { name }`, xlcalc's `BookFile`; a browser does not tell the folder), and with `session: { file }` `CELL("filename")` gives `[Budget.xlsx]Sheet1` and `INFO` knows the file, where without it they are those of a workbook not saved yet. The sheet shows no list of the session's diagnostics, so `CellIssue.fromFile` reaches the host only through xlcalc.

### Links to other workbooks

A formula that reads another workbook (`='[Prices.xlsx]Sheet1'!A1`) calculates over the values the workbook stores of it, as Excel calculates a file it opens without updating its links; the other file is never read. A host that has read the other workbook gives the links their values with `setExternalBooks(books)` (xlcalc 0.62.0's, every link at once as `BookInput.externalBooks` lists them): a new link after the others, more values of one, values read again. The workbook is calculated again in one step of undo, the changed cells reach the sheet as an edit's, and typed text then finds a new link by its file name. `saveXlsx` with `over` writes a new link as a part of the file and the values of a changed one into its part; removing a link, or moving it among the others, is refused (`KeptRefusal` with `code: 'links'`), since the file's formulas read a link by its number.

```ts
const links = (await source.book()).externalBooks ?? []
// the values the host read from Prices.xlsx, after the workbook's other links
await source.setExternalBooks([...links, { target: 'Prices.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 42, B2: 'kg' } }] }])
await source.edit([{ sheet: 'Plan', cell: 'C3', text: "='[Prices.xlsx]Sheet1'!A1*2" }]) // 84
```

### Calls outside the workbook

A host answers the calls of its own functions and of the functions of connections (`CUBEVALUE` and the other cube functions, `RTD`, `WEBSERVICE`, `STOCKHISTORY` and the rest) with xlcalc 0.68.0's `externals.resolve`, passed in `session`. An answer may be a promise. Each call of the session waits for the answers on their way within its `timeoutMs`, and without one until they arrive. Past it the call ends with its cells waiting: they show `#BUSY!` (a cube function's `#GETTING_DATA`), and the source takes each answer as it arrives, after its pending changes, through the session's `settleConnections`, sending the cells it changed as a `changed` event, so the sheet shows the answers without the user doing anything. Closing the source ends the waits: an answer still on its way is taken as none, a call it would lead to is not asked of the host, and nothing is sent after `close()`.

The deadline of every call is the calculator's `timeoutMs`: `pool: { timeoutMs }` for the pool the source starts, or the host's own calculator in `calc` (`createInlineCalc({ timeoutMs })`). `session: { timeoutMs }` bounds only the opening of the workbook; the edits after it then wait for their answers as long as they take.

```ts
const source = localSource(book, {
  pool: { timeoutMs: 200 }, // each call waits 200 ms for the answers, then shows #BUSY! and goes on
  session: { externals: { functions: ['BALANCE'], resolve: ({ name, args, signal }) => (name === 'BALANCE' ? fetchBalance(args[0], signal) : undefined) } },
})
```

A call's `timeoutMs` is also its calculation's deadline (xlcalc's `soft_timeout`), not counting the wait: a host gives one that its largest calculation fits in. When taking an answer runs past it, the source takes it once more with a deadline of a minute. A known gap of xlcalc 0.71.2: a stopped call can leave the session half calculated, some cells reading the answer and others still `#BUSY!`, and no later call completes it; give a `timeoutMs` that the calculation after an answer fits in. A failure there, or any other, is written to the console (`console.error`) and the sheet reads its windows again, so it shows what the session holds.

Refresh All (`refreshConnections`, Ctrl+Alt+F5) asks every call again, or with `names` those of the functions named: the source has it only with `externals.resolve`, as values given in advance (`externals.values`) answer nothing anew, and the formulas keeping their file's values would lose them to the answers of a function without its connection. It is not a step of undo ([calculation.md](calculation.md#refresh-all)).

In Manual only the cells that asked show their answers: a cell whose own formula calls the function, shown waiting for it since a formula entered or Calculate Now calculated it, takes the answer, as it ends that calculation. Every other cell waits for a calculation as after an edit: one that shows `#BUSY!` read from a cell that asked, one that reads the answer, and one whose waiting value is held while the sheet shows its value of the last calculation ([calculation.md](calculation.md#manual)). Not measured in Excel.

### CSV files

`openCsv` reads a CSV file into a workbook of one sheet, and `saveCsv` writes a sheet of the source's workbook as a CSV file, through xlcalc's `xlcalc/csv` (0.68.0, its `readCsv` and `writeCsv`), which is loaded only when one of them is called. A CSV file stores no language, so the host gives `locale`, the workbook's: each field is read as text typed into a cell in it, as Excel opens the file, so `1234,5` and `15.03.2026` are a number and a date in ru-RU, and the workbook takes the locale. A tag is loaded on demand, as for `openXlsx`. The result's `delimiter` and `encoding` say how the file was read. The encoding is found as Excel for Windows finds it: a byte order mark, UTF-8, or the locale's Windows code page. Fields are split at the locale's list separator, as Excel splits them (`,` in en-US, `;` in ru-RU: a ru-RU file opened in en-US is one column, as in Excel); a first line `sep=;` names the separator unless the host gives `delimiter`, a `.tsv` name reads at tabs, and `delimiter` names another (xlcalc's guess from the first lines is not used: Excel does not guess). A `File` names the sheet after itself, without the extension, as Excel names it.

One difference from Excel is deliberate: a field Excel takes as a formula (`=`, a `+` or `-` that begins no number, `@` before a call) stays text, and `issues` lists those fields with code `formula`. Excel would make them formulas, and a formula in a file from elsewhere could call the host's functions and the functions of connections (`WEBSERVICE`, `CUBEVALUE`, `RTD`), which xlcalc answers through the host's `externals.resolve`: opening a file would send requests the user did not ask for (CSV injection). `formulas: true` reads them as Excel does, in the locale's formula text, for files the host trusts. The other options are `readCsv`'s (`codesAsText`, `encoding`, `limits`), passed through.

`saveCsv` writes what Excel's CSV UTF-8 writes: each cell's value as it is displayed in the workbook's locale, at its list separator, with the byte order mark, formulas by their values. It writes one sheet, the workbook's active one unless `sheet` names another (the sheet's `state.sheet` is the one on screen), and reports the others in `issues`, as a CSV file holds one sheet. In Manual it calculates first, as `saveXlsx` does. `encoding: 'ansi'` writes Excel's plain CSV in the locale's Windows code page.

```ts
import { localSource, openCsv, saveCsv } from '@xlcalc/vue/local'

export async function openPrices(file: File): Promise<Uint8Array> {
  const { book, issues } = await openCsv(file, { locale: 'ru-RU' }) // Prices.csv: a sheet named Prices
  if (issues.length) console.warn(issues.map((i) => i.message))
  const source = localSource(book)
  try {
    await source.edit([{ sheet: book.sheets[0]!.name, cell: 'D2', text: '=B2*C2' }])
    return (await saveCsv(source)).data // UTF-8 with its mark, `;` between fields, `1234,5`
  } finally {
    await source.close()
  }
}
```

## Threads, the workbook and events

The pool starts its thread as `new Worker(new URL('./thread.js', import.meta.url))`, which Vite and other bundlers recognize; where a bundler serves the thread script from elsewhere, pass `pool: { spawn }`. Besides the `SheetSource` methods, the source has `ready`, `locale`, `history` (the steps `undo` and `redo` can take, for a toolbar), `book()`, `snapshot()` (the workbook and its values captured at one moment, which is what `writeXlsx` takes; it replaces `result()` of 0.2) and `close()`. It notifies its subscribers of every change it makes: `changed` for edits and widths, `reload` for structure, number formats, rows and columns, frozen panes and their undo, `repaint` with the formatted ranges for a style set by `format` (a style changes no text or size), and `sheets` when an edit changes the range a sheet's AutoFilter acts on ([autofilter.md](autofilter.md#data-typed-under-it)). The values the source caches for the display filter are kept through a style change, a freeze, and an undo or redo of an edit, since these change no cells or list the cells they change. An edit, a paste, a fill, an undo or a redo in a workbook with conditional formats sends `repaint` after its `changed`, with the ranges of the workbook's rules returned by the session: their cells may look different anywhere the rules apply (a color scale changes colors when a single number changes), and a cell update does not carry formatting. The values the source caches for search and the filter are kept then, unless a rule displays numbers with its own number format, and a display filter keeps its rows.
