# API reference

Everything the package exports, with signatures. What the sheet does with each piece, key by key, is in the other files of `docs/`; this one says what to pass and what comes back. The types are declared in `src/source.ts` (the source contract), `src/core/*.ts` (the sheet's state, events and methods), `src/messages.ts`, `src/memory-source.ts` and `src/local/`.

## Entries

| Import | What it gives | Runtime dependencies |
| --- | --- | --- |
| `xlcalc-sheet` | `SheetView`, `SheetToolbar`, `FormatDialog`, `memorySource`, `enMessages`, `ruMessages` and the types below | `vue` |
| `xlcalc-sheet/local` | `localSource`, `openXlsx`, `saveXlsx`, `groupIssues` | `vue`, `xlcalc` 0.43.0 or later (optional peer) |
| `xlcalc-sheet/style.css` | The look of the sheet, the toolbar and the dialog | |

The main entry bundles xlcalc's engine-free subpaths `xlcalc/a1` and `xlcalc/text` (addresses, the references of formula text, a paste's rectangle), and a book in ru-RU loads ru-RU's formula text as a chunk of its own.

Types of cells and styles come from `xlcalc/contract` (`CellValue`, `CellStyle`, `StylePatch`, `NumberFormat`, `ConditionalAppearance`, `StructureOp`). A host whose source is on a server needs `xlcalc` only in `devDependencies`, for these types; the main entry imports nothing of it at runtime.

## `SheetView`

The sheet: a windowed grid with the formula bar, tabs, menus, Find, the precedents' panel and the dialogs. It reads its data only through `source`.

### Props (`SheetViewProps`)

| Prop | Type | Default | Meaning |
| --- | --- | --- | --- |
| `source` | `SheetSource` | required | Where the sheet takes its data from. A new source is a new book. |
| `messages` | `Partial<SheetMessages>` | English | Interface texts; a part overrides the English defaults. |
| `storageKey` | `string` | none | Key in the browser's storage for the freeze choice, widths and the recently used functions of Insert Function; without it they live with the component. |
| `focus` | `{ ref: string } \| null` | `null` | A place to show: `Sheet!B12:D15`, `B7`, `Sheet`. A new object is a new jump. |
| `cellUrl` | `(ref: string) => string` | none | The host's page address for a place; without it there is no "Link to the cell" button. |
| `readOnly` | `boolean` | `false` | No editing whatever the source offers. A source without `edit` is read-only anyway. |
| `formulaBar` | `boolean` | `true` | The formula bar over the grid. |
| `locale` | `string` | none | The book's locale for numbers the component writes itself (a number's edit text, selection totals); `localSource(...).locale` for a local book. |
| `hideLoneTab` | `boolean` | `false` | No tabs for a book of one sheet. |
| `edgeRestoresWidth` | `boolean` | `false` | A double click on a column's edge brings the book's width back instead of fitting the column. |
| `sort` | `ListSort \| null` | `null` | List mode: the order the host sorted the list in, an arrow on that column's title. |
| `openRows` | `boolean` | `false` | List mode: the first column opens its row, and Enter on a cell that is not a drill-down does too. |

### Events

| Event | Payload | When |
| --- | --- | --- |
| `state` | `SheetViewState \| null` | The selection, the active cell or its look changed; `null` while there is no sheet. A toolbar and `FormatDialog` take it as their `state`. |
| `openLink` | `SheetLinkEvent` | A cell's link that is not a place in the book was clicked, and the host listens: the host opens it (its `href`, or a `target` of its own). Without a listener a web link opens as an ordinary link. A link into the book (`location`) the sheet follows itself. |
| `sort` | `ListSort` | List mode: a sortable column's title was clicked. The host sorts the list, sets `sort` and sends `reload`. |
| `cellDrill` | `{ sheet: string; cell: string }` | List mode: a drill-down cell (`SheetWindow.drill`) was clicked or took Enter. |
| `rowOpen` | `{ sheet: string; row: number }` | List mode, with `openRows`: a row was opened from its first column or by Enter. |

### Slot

| Slot | Scope | Meaning |
| --- | --- | --- |
| `fallback` | `{ failure: SheetFailure; retry: () => void }` | The book cannot be shown: the source's error and a retry. Without the slot the sheet shows its own message and a Retry button. |

### `SheetViewState`

| Field | Type | Meaning |
| --- | --- | --- |
| `sheet` | `string` | The sheet shown. |
| `ref` | `string` | The selection as Excel writes it: `B2:D5`, `B:D`, `2:5`, `1:1048576`. |
| `formula` | `string?` | The active cell's formula as shown. |
| `value` | `string?` | The active cell's shown text. |
| `style` | `CellStyle?` | The style the active cell shows; absent while its window is not loaded. |
| `numberFormat` | `string?` | The active cell's number format code; absent for General. |
| `merged` | `string?` | The merged area the active cell is in. |
| `canMerge` | `true?` | The merge commands act: the source has `structure`, the sheet is editable, no display filter. |

`SheetLinkEvent` is a window's link (`href`, `location`, `tooltip`, `target`) with the `sheet` and `cell` it is on. `SheetFailure` is `{ error: unknown; text: string }`.

### `SheetViewApi`

`SheetView` exposes these methods through a template ref (`ref="sheet"`, `const sheet = ref<SheetViewApi | null>(null)`). Each goes through the source as the sheet's own keys go: the outcome is applied, its selection set without the flash and the step back that `focus` makes, a refusal shown in the sheet's message line, and after `format`, `structure`, `outline` or `freezePanes` the sheet reads itself again unless the source sent `reload`. A method that returns a promise of `EditOutcome | null` gives the source's outcome, or `null` when nothing was sent (no such method on the source, a read-only sheet, an undo still out).

| Method | Returns | Meaning |
| --- | --- | --- |
| `undo()` | `Promise<EditOutcome \| null>` | The source's `undo`. |
| `redo()` | `Promise<EditOutcome \| null>` | The source's `redo`. |
| `format(patches: SheetFormatPatch[])` | `Promise<EditOutcome \| null>` | Formats the active sheet in one call of the source's `format`. A patch without `range` goes over the selection as Excel writes it (`B:B`, `2:5`). |
| `structure(op: SheetOp)` | `Promise<EditOutcome \| null>` | A structure operation; the sheet shown after it is the one Excel shows. |
| `outline(change: OutlineChange)` | `Promise<EditOutcome \| null>` | Group, Ungroup, a group's button, a level button. |
| `freezePanes(mode?: FreezeMode)` | `Promise<EditOutcome \| null>` | Excel's Freeze Panes through the source's `freeze`: `'panes'` (the default, at the active cell), `'unfreeze'`, `'topRow'`, `'firstColumn'`. Nothing is sent when there is nothing to freeze or unfreeze. |
| `merge(command: MergeCommand)` | `Promise<EditOutcome \| null>` | `'mergeCenter'`, `'mergeAcross'`, `'mergeCells'` or `'unmergeCells'` over the selection through `structure`, with Excel's warning first when values would be lost. `null` also for one cell outside a merged area, under a display filter and on the warning's Cancel. |
| `formatCells()` | `boolean` | Opens Excel's Format Cells dialog; `false` when the sheet cannot format. |
| `reapplyFilter()` | `Promise<boolean>` | Excel's Reapply; `false` with no display filter in force. |
| `explain()` | `boolean` | Opens the precedents' panel at the active cell; `false` without the source's `explain`. |

`SheetFormatPatch` is a `FormatPatch` whose `range` may be left out. `FreezeMode` is `'panes' | 'unfreeze' | 'topRow' | 'firstColumn'`, `MergeCommand` is `'mergeCenter' | 'mergeAcross' | 'mergeCells' | 'unmergeCells'`. `ListSort` is `{ col: number; dir: 'asc' | 'desc' }`, the column 1-based.

## `SheetToolbar`

Excel's formatting row over the same source; shown only when the source has `format`. See [toolbar.md](toolbar.md).

| Prop | Type | Meaning |
| --- | --- | --- |
| `source` | `ToolbarSource` (the same as `SheetSource`) | The sheet's source. |
| `state` | `SheetViewState \| null` | What the sheet's `state` event gave; `null` leaves nothing to press. |
| `messages` | `Partial<SheetMessages>` | Interface texts. |
| `locale` | `string` | The book's locale: the Number Format box names its categories and items in it; en-US without it. |
| `sheet` | `SheetViewApi \| null` | The sheet's template ref. With it the toolbar formats, undoes, redoes and merges through the sheet; without it, through the source, and it emits `select` and `error`. |

| Event | Payload | When |
| --- | --- | --- |
| `select` | `{ sheet: string; range: string }` | Where the selection goes after an undo or a redo, for the sheet's `focus`; not with `sheet`. |
| `error` | `string` | The source refused or failed a call; not with `sheet`. |
| `done` | none | A call is done: the host may give the grid its focus back. |
| `formatCells` | none | Format Cells… was pressed. With a listener the button opens the host's dialog; without one it opens the sheet's (`sheet`), and without both there is no button. |

## `FormatDialog`

Excel's Format Cells over the selection, for a host that places its own dialog. `SheetView` has one of its own; see [toolbar.md](toolbar.md#formatdialog).

| Prop | Type | Meaning |
| --- | --- | --- |
| `open` | `boolean` | Shown while true; each opening starts again from the active cell's look. |
| `state` | `SheetViewState \| null` | What the sheet's `state` event gave; `null` shows nothing. |
| `messages` | `Partial<SheetMessages>` | Interface texts. |
| `locale` | `string` | The book's locale for the Number tab. |

| Event | Payload | When |
| --- | --- | --- |
| `apply` | `SheetFormatPatch[]` | OK with changes: patches for the sheet's `format`. |
| `close` | none | OK or Cancel: the host hides the dialog. |

## The source: `SheetSource`

The component knows its data only through this interface. Rows and columns are 1-based, cells are A1 addresses, formulas are shown and typed in the source's own locale. Two methods are required; each optional one turns a feature on. An editing method answers with an `EditOutcome`.

### Required

| Method | Meaning |
| --- | --- |
| `sheets(signal?) => Promise<SheetInfo[]>` | The book's sheets in order; read again after a structure edit and on `reload`. |
| `window(req: WindowRequest, signal?) => Promise<SheetWindow \| SheetPending>` | A rectangle of a sheet, at most `maxCells` cells. `SheetPending` makes the sheet wait and ask again. |

### Optional methods

| Method | Turns on | Without it | More |
| --- | --- | --- | --- |
| `edit(entries: CellEntry[], opts?: EditOptions)` | Editing. `text: null` clears. | The sheet is read-only. | [editing.md](editing.md) |
| `paste(sheet, cell, text)` | A paste read by the source, one step of undo. | Fields go to `edit`. | [editing.md](editing.md) |
| `fill(sheet, range, text, active, opts?)` | Control+Return over a selection. | Only the active cell takes the text. | [editing.md](editing.md) |
| `undo()`, `redo()` | Undo and redo. | No undo. | [editing.md](editing.md) |
| `structure(op: SheetOp)` | Insert, delete, rename, move, hide sheets and lines; copy and move ranges; merge. | No structure menu, no fill handle, a copy pastes as text. | [sheets.md](sheets.md), [clipboard.md](clipboard.md) |
| `layout(change: LayoutChange)` | Widths, heights, hiding and fitting lines. | A width lives until the page reloads. | [sheets.md](sheets.md) |
| `copyText(sheet, range, opts?) => Promise<string>` | Excel's clipboard text of a range. | The sheet copies the texts its windows show. | [clipboard.md](clipboard.md) |
| `format(sheet, patches: FormatPatch[])` | Formatting, the toolbar, Format Cells, Control+B/I/U. | No toolbar. | [toolbar.md](toolbar.md) |
| `subscribe(listener: (e: SourceEvent) => void) => () => void` | Changes made elsewhere. Returns the unsubscribe function. | The sheet sees only its own edits. | [editing.md](editing.md) |
| `validation(sheet, cell, signal?) => Promise<SheetValidation \| null>` | Excel's dropdown and input message. | None. | [validation.md](validation.md) |
| `pick(sheet, cell, item, text)` | A picked list item written as Excel writes it. | The item goes to `edit` with `accept`. | [validation.md](validation.md) |
| `freeze(sheet, panes: { rows; cols } \| null)` | Freeze Panes in the menu. | The lock alone hides a book's freeze. | [view.md](view.md) |
| `functionNames(signal?) => Promise<string[]>` | Formula AutoComplete and Insert Function. | No list, the fx button off. | [autocomplete.md](autocomplete.md) |
| `functionSignature(name, signal?) => Promise<SheetFunctionSignature \| null>` | A function's arguments for Excel's tip while a formula is typed (`SUM(number1, [number2], ...)`), and in Insert Function its syntax, `description` and `category` with `categoryName` (xlcalc 0.42); the name as typed, any case; null for a function the source does not know. | No tip; Insert Function lists the names alone. | [autocomplete.md](autocomplete.md) |
| `explain(sheet, cell) => Promise<string>` | Trace Precedents. | No panel. | [view.md](view.md) |
| `end(sheet, cell, direction, signal?) => Promise<string>` | Control and an arrow as Excel's Range.End. | The edge of the data the sheet knows. | [editing.md](editing.md) |
| `stats(sheet, range, filter?, signal?) => Promise<RangeStats>` | Selection totals over the whole range. | Counted over the loaded windows. | |
| `find(req: FindRequest, signal?) => Promise<FindResult>` | Find. | No search. | [find-replace.md](find-replace.md) |
| `replace(req: FindRequest, by) => Promise<ReplaceOutcome>` | Replace. | No replace row. | [find-replace.md](find-replace.md) |
| `outline(change: OutlineChange)` | The outline's buttons, Group and Ungroup. | The bars are drawn and do not press. | [sheets.md](sheets.md) |
| `visibleRows(sheet, filter, signal?) => Promise<{ rows: number[] \| null; total }>` | The display filter and sort. | No filter. | [view.md](view.md) |
| `columnValues(sheet, col, filter, signal?) => Promise<{ values: { text; count }[]; more }>` | The filter funnel's list of values. | No list in the funnel. | |

| Property | Type | Meaning |
| --- | --- | --- |
| `maxCells` | `number` | The cell limit of one window, 20,000 by default. |
| `history` | `{ undo: number; redo: number }` | The steps `undo` and `redo` can take, for the toolbar's buttons. |

### Windows and sheets

`WindowRequest` is `{ sheet, row, rows, col, cols, filter? }`; under a display filter `row` is a place among the visible rows.

`SheetWindow`:

| Field | Type | Meaning |
| --- | --- | --- |
| `row`, `col` | `number` | The window's top-left cell. |
| `text` | `string[][]` | Shown text, rows × columns; an empty cell is `''`. |
| `values` | `CellValue[][]` | Values of the same cells, for totals and alignment. |
| `formulas` | `Record<string, string>` | Formulas by address as shown, without `=`. |
| `editText` | `Record<string, string>?` | The text each cell opens with; a cell left out opens empty. Without it the component writes it from the value and formula. |
| `colors` | `Record<string, string>` | The colour of the number format's section, `#RRGGBB`. |
| `styleIds`, `styles` | `Record<string, number>`, `CellStyle[]` | Each cell's style by index into `styles`. |
| `numberFormats` | `Record<string, string>?` | The format code each cell shows by; a cell left out is General. |
| `widths` | `(number \| null)[]` | Column widths in xlcalc's units, `null` for the default. |
| `rowHeights` | `Record<number, number>` | Row heights in points. |
| `defaultRowHeight` | `number?` | In points. |
| `merges` | `string[]` | Merged areas. |
| `tables` | `{ name; range; headerRow }[]?` | Tables, whose header and border are drawn. |
| `appearance` | `Record<string, ConditionalAppearance>?` | What conditional formats give cells, as xlcalc computes it. |
| `bars` | `Record<string, { ratio; color }>?` | Old-style data bars; `appearance` takes their place. |
| `links` | `Record<string, { href?; location?; tooltip?; target? }>?` | Hyperlinks. |
| `notes` | `Record<string, { author?; text }>?` | Notes. |
| `notice` | `string?` | A notice above the grid. |
| `drill` | `Record<string, true>?` | List mode: drill-down cells, underlined; a click or Enter emits `cellDrill`. An empty cell is no drill-down. |

`SheetPending` is `{ pending: { elapsedMs } }`: the book is recomputing, and the sheet asks again.

`SheetInfo`:

| Field | Type | Meaning |
| --- | --- | --- |
| `name` | `string` | The sheet's name. |
| `maxRow`, `maxCol` | `number` | The extent of the data. |
| `freeze` | `{ rows; cols }?` | The book's frozen panes. |
| `hidden` | `boolean \| 'very'?` | Not on the tabs; `'very'` is not offered to unhide either. |
| `locked` | `{ range; reason }[]?` | Rectangles that cannot be edited, with the reason shown. |
| `columns` | `{ title; align?; sortable? }[]?` | List mode: titles in place of column letters, aligned by `align` (the totals too); a `sortable` title's click emits `sort`. |
| `rowNumbers` | `boolean?` | Row numbers on the left, on by default. |
| `totals` | `(string \| null)[]?` | List mode: a totals row held at the bottom of the frame, the shown text by the sheet's columns. |
| `outline` | `SheetOutline?` | Excel's outline: `rows` and `columns` as `OutlineGroup[]` (`{ from, to, level, collapsed }`), `summaryBelow`, `summaryRight`. |

### Edits and outcomes

`CellEntry` is `{ sheet, cell, text: string | null }`. `EditOptions` is `{ accept?: true; inPlace?: true }`: `accept` writes a value a validation rule refused after the person kept it, `inPlace` says the text was edited from the cell's own (F2, a double click, the formula bar), so Excel's percent editor adds no `%`.

`EditOutcome` is one of:

- `{ ok: true; changes: CellUpdate[]; select?: { sheet; range } }`: `select` is where the selection goes after `paste`, `fill`, `undo` or `redo`.
- `{ ok: false; reason; message; cells?; at?; alert? }`, where `reason` is `'conflict'` (the book changed elsewhere), `'refused'` (locked or not allowed), `'invalid'` (the formula does not parse; `at` is the position in the typed text), `'busy'` (try again later) or `'failed'`. `alert` (`SheetValidationAlert`) makes the sheet show Excel's validation alert.

`CellUpdate` is `{ sheet, cell, text, value, color?, formula?, editText?, numberFormat? }`; `formula`, `editText` and `numberFormat` come only for edited cells.

`SourceEvent`, sent to `subscribe`'s listener:

- `{ kind: 'changed'; changes: CellUpdate[] }`: the loaded windows are patched.
- `{ kind: 'reload'; message? }`: the sheets and windows are read again.
- `{ kind: 'repaint'; ranges: { sheet; range }[] }`: the windows that meet the ranges are read again in place.

`FormatPatch` is `{ range, style?: StylePatch | null, numberFormat?: NumberFormat | null }`; `style: null` removes the cells' own style. `LayoutChange` is `{ sheet, columns, width?, hidden?, fit? }` or `{ sheet, rows, height?, hidden?, fit? }`. `SheetOp` is xlcalc's `StructureOp` without its table operations. `OutlineChange` is `{ op: 'group' | 'ungroup', sheet, rows | columns }`, `{ op: 'collapse' | 'expand', sheet, rows | columns, level }` or `{ op: 'showLevel', sheet, axis, level }`.

### Find, filter, validation

| Type | Shape |
| --- | --- |
| `FindRequest` | `{ sheet, query, range?, matchCase?, wholeCell?, lookIn?: 'formulas' \| 'values', byColumns?, after?, filter? }`. `query` takes Excel's wildcards `*`, `?` and `~`. |
| `FindResult` | `{ hits: { row, col, text }[]; total }` |
| `ReplaceOutcome` | `EditOutcome & { replaced: number; refused?: { sheet, cell, message }[] }` |
| `SheetFilter` | `{ head: number; cols?: Record<number, { values?: string[]; text?: string }>; sort?: { col; dir: 'asc' \| 'desc' } }` |
| `RangeStats` | `{ cells, numbers, sum, min, max }` |
| `SheetFunctionSignature` | xlcalc's `FunctionSignature` with `description?`, one line of what the function does in the source's locale, for Insert Function. |
| `SheetValidation` | `{ list?: string[]; prompt?: { title?; text } }` |
| `SheetValidationAlert` | `{ severity: 'stop' \| 'warning' \| 'information'; title?; message; sheet; cell }` |

## Interface texts: `SheetMessages`

`enMessages` (the default) and `ruMessages` are full sets. A host passes a set or a part of one as `messages`; the keys and their arguments are in `src/messages.ts`.

```ts
import { ruMessages } from 'xlcalc-sheet'
const messages = { ...ruMessages, loading: 'Загружаю…' }
```

## `memorySource`

`memorySource(book: MemoryBook, options?: MemorySourceOptions): MemorySource` is a source over a book held in memory, for an example or a test. It has no engine: a formula's value is whatever the cell's `value` says.

| Type | Shape |
| --- | --- |
| `MemoryBook` | `{ sheets: MemorySheet[]; styles?: CellStyle[]; functions?: string[] }`; `functions` turns on `functionNames`. |
| `MemorySheet` | `SheetInfo` without `maxRow`, `maxCol` and `outline`, plus `cells: Record<string, MemoryCell \| CellValue>`, `widths?` (by `B` or `B:D`, xlsx units, 0 hides), `heights?` (by `5` or `5:9`, points, 0 hides), `defaultRowHeight?`, `merges?`, `tables?`, `notice?`, `outline?` (levels by line or span, `collapsed`, `summaryBelow`, `summaryRight`), `validations?: MemoryValidation[]`. |
| `MemoryCell` | `{ value?, text?, formula?, style?, color?, link?, note?, bar?, appearance?, drill? }` |
| `MemoryValidation` | `{ range; list?: string[]; whole?: { min?; max? }; prompt?: { title?; text }; alert?: { severity?; title?; message? } }` |
| `MemorySourceOptions` | `{ delay?: number; maxCells?: number }`; `delay` makes every answer take that many milliseconds. |

`MemorySource` is a `SheetSource` with four more methods for tests and examples: `failNext(...outcomes)` refuses the next edits with these outcomes, `emit(event)` sends an event to the subscribers, `setCells(sheet, entries)` changes cells as another tab would, `recompute(ms)` makes windows answer `pending` for that long.

## `xlcalc-sheet/local`

A source over an xlcalc book computed in the browser, and xlsx in and out. See [local.md](local.md).

### `localSource(input: BookInput, opts?: LocalSourceOptions): LocalSource`

The options, `LocalSourceOptions`:

| Option | Type | Meaning |
| --- | --- | --- |
| `calc` | `'pool' \| 'inline' \| Calc` | Where the book computes: Web Workers (the default where `Worker` exists), the page's thread in portions (the default elsewhere), or the host's calculator, which the source does not close. |
| `pool` | `CalcPoolOptions` | Options of the pool the source starts, such as `spawn` for a bundler that moves the thread's script. |
| `session` | `AsyncSessionOptions` | Options of the book's session; `{ cached }` from `openXlsx` holds a file's stored values. |
| `maxCells` | `number` | The cell limit of one window, 20,000 by default. |
| `notice` | `string` | A notice every window carries. |

`LocalSource` has every method of `SheetSource` (the type marks `copyText` optional, the source has it), and:

| Member | Type | Meaning |
| --- | --- | --- |
| `locale` | `string` | The book's locale tag; pass it to `SheetView`'s `locale`. |
| `ready` | `Promise<void>` | Settles when the book is computed; rejects when it could not be opened. |
| `history` | `SessionHistory` | The steps `undo` and `redo` can take now. |
| `book()` | `Promise<BookInput>` | The book with every change so far. |
| `snapshot()` | `Promise<BookSnapshot>` | The book and its values taken at one moment, as `writeXlsx` takes them. |
| `close()` | `Promise<void>` | Closes the session, and the pool when the source started it. |

### Files

| Function | Returns | Meaning |
| --- | --- | --- |
| `openXlsx(data: Blob \| ArrayBuffer \| Uint8Array, opts?: XlsxIssuesOptions)` | `Promise<XlsxRead>` (`{ book, cached, issues }`) | Reads a file into a book and the values it stores. |
| `saveXlsx(source: LocalSource, opts?: XlsxIssuesOptions & { mode?: 'formulas' \| 'values' })` | `Promise<XlsxWritten>` (`{ data, issues }`) | Writes the source's book with its values; `values` writes no formulas. |
| `groupIssues(issues: XlsxIssue[])` | `IssueGroup[]` | Issues by kind, in the order their kinds first appear, for a notice of what a file lost. |

`XlsxIssuesOptions` is `{ locale?: string | Locale }`: the language of the issues' messages, English without it. `IssueGroup` is `{ code, count, sheets: string[], message }`, `message` being the first issue's.

## CSS variables

`xlcalc-sheet/style.css` styles `.xs-sheet`, `.xs-toolbar` and the dialog by `--xs-*` variables; a host restyles or darkens the sheet by overriding them on `.xs-sheet` (and `.xs-toolbar`) or an ancestor.

| Variable | Default | What |
| --- | --- | --- |
| `--xs-bg`, `--xs-fg`, `--xs-muted` | `#ffffff`, `#1f2328`, `#6b7280` | Background, text, secondary text. |
| `--xs-font`, `--xs-font-size` | `system-ui, …`, `12.5px` | The interface's font. |
| `--xs-grid` | `#e1e3e6` | Grid lines and the interface's borders. |
| `--xs-head-bg`, `--xs-head-fg` | `#f3f4f6`, `#5f6368` | Row numbers and column letters. |
| `--xs-head-on-bg`, `--xs-head-on-fg` | `#d3e3fd`, `#0b57d0` | Headers of selected lines. |
| `--xs-accent`, `--xs-sel-tint` | `#1a73e8`, `rgb(26 115 232 / 0.12)` | The selection's border and tint. |
| `--xs-link`, `--xs-note` | `#1558d6`, `#dc2626` | Hyperlinks, a note's mark. |
| `--xs-flash`, `--xs-hit` | `rgb(250 204 21 / 0.85)`, `rgb(250 204 21 / 0.45)` | A jump's flash, Find's hits. |
| `--xs-freeze-line`, `--xs-border-auto` | `#9aa0a6`, `#000000` | The frozen panes' line, a border of automatic colour. |
| `--xs-table-head-bg`, `--xs-table-line` | `#e8eef8`, `#8faadc` | Tables. |
| `--xs-locked`, `--xs-pending`, `--xs-error` | `#8b8f96`, `#8b8f96`, `#b42318` | Locked cells, pending edits, errors. |
| `--xs-bar-bg` | `#fafafa` | The formula bar, the toolbar and the precedents' panel. |
| `--xs-message-bg`, `--xs-message-fg` | `#1f2328`, `#ffffff` | The message line. |
| `--xs-copy-line`, `--xs-fill-frame` | `#1a73e8`, `#5f6368` | A copy's moving border, the fill handle's frame. |
| `--xs-outline-bg`, `--xs-outline-line` | `#fafafa`, `#8c8c8c` | The outline's bars and brackets. |
| `--xs-outline-button`, `--xs-outline-tick` | `13px`, `5px` | The outline's button size and a bracket's tick. |
| `--xs-outline-button-bg`, `--xs-outline-button-fg`, `--xs-outline-button-border` | `#ffffff`, `#3c4043`, `#9aa0a6` | The outline's buttons. |
| `--xs-explain-w`, `--xs-explain-h` | `340px`, `45%` | The precedents' panel: its width, and its height on a phone. |
| `--xs-prompt-bg`, `--xs-prompt-fg`, `--xs-prompt-line` | `#ffffe1`, `#1f2328`, `#767676` | Data validation's input message. |
| `--xs-alert-stop`, `--xs-alert-warning`, `--xs-alert-information` | `#d13438`, `#ffb900`, `#0f6cbd` | The validation alert's icons. |
| `--xs-fl-line`, `--xs-fl-on-fg` | `#7a7a7a`, `#ffffff` | Formula AutoComplete's list. |
