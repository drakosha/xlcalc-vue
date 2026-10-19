# API reference

Everything the package exports, with signatures. How the sheet uses each part, key by key, is described in the other files of `docs/`; this page covers what to pass and what is returned. The types are declared in `src/contract/`, behind `src/contract.ts`: `source.ts` (the source contract), `view.ts` (the sheet's props, state, events and methods), `memory.ts` (`memorySource`) and `messages.ts` (the interface texts' keys); those of `@xlcalc/vue/local` are in `src/local/types.ts`.

## Entry points

| Import | Exports | Runtime dependencies |
| --- | --- | --- |
| `@xlcalc/vue` | `SheetView`, `SheetToolbar`, `FormatDialog`, `memorySource`, `enMessages`, `ruMessages` and the types below | `vue` |
| `@xlcalc/vue/local` | `localSource`, `openXlsx`, `saveXlsx`, `groupIssues` | `vue`, `xlcalc` 0.46.0 or later (optional peer) |
| `@xlcalc/vue/style.css` | Styles of the sheet, the toolbar and the dialog | |

The main entry point bundles xlcalc's engine-free subpaths `xlcalc/a1` and `xlcalc/text` (addresses, references in formula text, the range of a paste), and a workbook in ru-RU loads the ru-RU formula text as a separate chunk. With a source without `editText`, `xlcalc/entry` (parsing of typed entries, without the engine) is loaded as another chunk, bundled in full separately from the main entry point, when the first cell needs its text; `memorySource` loads it for Find and Replace.

Cell and style types come from `xlcalc/contract` (`CellValue`, `CellStyle`, `StylePatch`, `NumberFormat`, `ConditionalAppearance`, `StructureOp`). A host whose source is on a server needs `xlcalc` only in `devDependencies`, for these types; the main entry point imports nothing from it at runtime.

## `SheetView`

The sheet: a virtualized grid with the formula bar, sheet tabs, context menus, Find, the Trace Precedents panel and the dialogs. It reads its data only through `source`.

```vue
<script setup lang="ts">
import { shallowRef } from 'vue'
import { SheetView, memorySource, type SheetLinkEvent, type SheetViewState } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({
  sheets: [
    {
      name: 'Summary',
      cells: {
        A1: 'Total',
        B1: { value: 60, formula: 'SUM(Data!B1:B3)' },
        A2: { value: 'The data', link: { location: 'Data!A1' } }, // a location in the workbook: the sheet goes there itself
        A3: { value: 'The report', link: { href: 'https://example.com/report' } }, // a web link: the host opens it
      },
    },
    { name: 'Data', cells: { A1: 'North', B1: 10, A2: 'South', B2: 20, A3: 'West', B3: 30 } },
  ],
})

const focus = shallowRef<{ ref: string } | null>(null)
const state = shallowRef<SheetViewState | null>(null)
// Each new object is a new jump, even to the same location
const show = (ref: string) => (focus.value = { ref })
// The host page's address for a location, for the cell menu's "Link to the cell"
const cellUrl = (ref: string) => `${location.origin}${location.pathname}#${encodeURIComponent(ref)}`
const openLink = (link: SheetLinkEvent) => window.open(link.href, '_blank', 'noopener')
</script>

<template>
  <button @click="show('Data!B1:B3')">Show the data</button>
  <span>{{ state ? `${state.sheet}!${state.ref}` : '' }}</span>
  <SheetView :source="source" :focus="focus" :cell-url="cellUrl" style="height: 360px" @state="state = $event" @open-link="openLink">
    <template #fallback="{ failure, retry }">
      <p>{{ failure.text }} <button @click="retry">Try again</button></p>
    </template>
  </SheetView>
</template>
```

### Props (`SheetViewProps`)

| Prop | Type | Default | Meaning |
| --- | --- | --- | --- |
| `source` | `SheetSource` | required | The sheet's data source. A new source means a new workbook. |
| `messages` | `Partial<SheetMessages>` | English | Interface texts; a partial set overrides the English defaults. |
| `storageKey` | `string` | none | Key in the browser's storage for the frozen panes display choice, column widths and Insert Function's recently used functions; without it they last as long as the component. |
| `focus` | `{ ref: string } \| null` | `null` | A location to show: `Sheet!B12:D15`, `B7`, `Sheet`. Each new object is a new jump. |
| `cellUrl` | `(ref: string) => string` | none | The host page's address for a location; without it there is no "Link to the cell" button. |
| `readOnly` | `boolean` | `false` | Disables editing regardless of what the source supports. A source without `edit` is read-only anyway. |
| `formulaBar` | `boolean` | `true` | Shows the formula bar above the grid. |
| `locale` | `string` | none | The workbook's locale for texts the component produces itself (a cell's edit text without `editText`, selection totals); `localSource(...).locale` for a local workbook. |
| `hideLoneTab` | `boolean` | `false` | Hides the tabs of a workbook with a single sheet. |
| `edgeRestoresWidth` | `boolean` | `false` | Double-clicking a column edge restores the workbook's width instead of fitting the column. |
| `sort` | `ListSort \| null` | `null` | List mode: the order the host sorted the list in, shown as an arrow on that column's title. |
| `openRows` | `boolean` | `false` | List mode: the first column opens its row, and so does Enter on a cell that is not a drill-down. |

### Events

| Event | Payload | When |
| --- | --- | --- |
| `state` | `SheetViewState \| null` | The selection, the active cell or its formatting changed; `null` while there is no sheet. A toolbar and `FormatDialog` take it as their `state`. |
| `openLink` | `SheetLinkEvent` | A cell's hyperlink that does not point into the workbook was clicked, and the host listens: the host opens it (its `href`, or a `target` of its own). Without a listener a web link opens as an ordinary link. The sheet follows links into the workbook (`location`) itself. |
| `sort` | `ListSort` | List mode: a sortable column's title was clicked. The host sorts the list, sets `sort` and sends `reload`. |
| `cellDrill` | `{ sheet: string; cell: string }` | List mode: a drill-down cell (`SheetWindow.drill`) was clicked or Enter was pressed on it. |
| `rowOpen` | `{ sheet: string; row: number }` | List mode, with `openRows`: a row was opened from its first column or by Enter. |

### Slot

| Slot | Scope | Meaning |
| --- | --- | --- |
| `fallback` | `{ failure: SheetFailure; retry: () => void }` | The workbook cannot be displayed: the source's error and a retry function. Without the slot the sheet shows its own message and a Retry button. |

### `SheetViewState`

| Field | Type | Meaning |
| --- | --- | --- |
| `sheet` | `string` | The displayed sheet. |
| `ref` | `string` | The selection as Excel writes it: `B2:D5`, `B:D`, `2:5`, `1:1048576`. |
| `formula` | `string?` | The active cell's formula as displayed. |
| `value` | `string?` | The active cell's displayed text. |
| `style` | `CellStyle?` | The active cell's style; absent while its window is not loaded. |
| `numberFormat` | `string?` | The active cell's number format code; absent for General. |
| `merged` | `string?` | The merged range containing the active cell. |
| `canMerge` | `true?` | The merge commands are available: the source has `structure`, the sheet is editable, and no display filter is on. |

`SheetLinkEvent` is a window's hyperlink (`href`, `location`, `tooltip`, `target`) with the `sheet` and `cell` it belongs to. `SheetFailure` is `{ error: unknown; text: string }`.

### `SheetViewApi`

`SheetView` exposes these methods through a template ref (`ref="sheet"`, `const sheet = ref<SheetViewApi | null>(null)`). Each call goes through the source just as the sheet's own keyboard commands do: the outcome is applied, its selection is set without the flash and the Back history entry that `focus` creates, a rejection is displayed in the sheet's message line, and after `format`, `structure`, `outline` or `freezePanes` the sheet reloads unless the source has sent `reload` (after `format`, also unless it has sent `repaint`, which re-reads the windows over its ranges). A method that returns a promise of `EditOutcome | null` resolves to the source's outcome, or to `null` when nothing was sent (the source lacks the method, the sheet is read-only, or an undo is still pending).

| Method | Returns | Meaning |
| --- | --- | --- |
| `undo()` | `Promise<EditOutcome \| null>` | The source's `undo`. |
| `redo()` | `Promise<EditOutcome \| null>` | The source's `redo`. |
| `format(patches: SheetFormatPatch[])` | `Promise<EditOutcome \| null>` | Formats the active sheet in one call to the source's `format`. A patch without `range` applies to the selection, written as Excel writes it (`B:B`, `2:5`). |
| `structure(op: SheetOp)` | `Promise<EditOutcome \| null>` | A structure operation; the sheet displayed afterwards is the one Excel displays. |
| `outline(change: OutlineChange)` | `Promise<EditOutcome \| null>` | Group, Ungroup, a group's button, a level button. |
| `freezePanes(mode?: FreezeMode)` | `Promise<EditOutcome \| null>` | Excel's Freeze Panes through the source's `freeze`: `'panes'` (the default, at the active cell), `'unfreeze'`, `'topRow'`, `'firstColumn'`. Nothing is sent when there is nothing to freeze or unfreeze. |
| `merge(command: MergeCommand)` | `Promise<EditOutcome \| null>` | `'mergeCenter'`, `'mergeAcross'`, `'mergeCells'` or `'unmergeCells'` over the selection through `structure`, with Excel's warning first when values would be lost. `null` also for a single cell outside a merged range, under a display filter, and on Cancel in the warning. |
| `formatCells()` | `boolean` | Opens Excel's Format Cells dialog; `false` when the sheet cannot be formatted. |
| `reapplyFilter()` | `Promise<boolean>` | Excel's Reapply; `false` when no display filter is on. |
| `explain()` | `boolean` | Opens the Trace Precedents panel for the active cell; `false` if the source has no `explain`. |

`SheetFormatPatch` is a `FormatPatch` whose `range` is optional. `FreezeMode` is `'panes' | 'unfreeze' | 'topRow' | 'firstColumn'`, `MergeCommand` is `'mergeCenter' | 'mergeAcross' | 'mergeCells' | 'unmergeCells'`. `ListSort` is `{ col: number; dir: 'asc' | 'desc' }`, with a 1-based column.

The host's own buttons calling the sheet's methods; the source here is `localSource`, which has `format`, `undo`, `structure` and `freeze`:

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { SheetView, type SheetViewApi } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({
  sheets: [
    {
      name: 'Plan',
      cells: {
        A1: { value: 'Item' }, B1: { value: 'Sum' },
        A2: { value: 'Rent' }, B2: { value: 1200 },
        A3: { value: 'Food' }, B3: { value: 400 },
        A4: { value: 'Total' }, B4: { formula: 'SUM(B2:B3)' },
      },
    },
  ],
})
onUnmounted(() => void source.close())

const sheet = ref<SheetViewApi | null>(null)
// Bold on a yellow fill over the selection, a single undo step
const highlight = () => sheet.value?.format([{ style: { font: { bold: true }, fill: { color: '#FFF2CC' } } }])
</script>

<template>
  <button @click="sheet?.undo()">Undo</button>
  <button @click="highlight">Highlight</button>
  <button @click="sheet?.freezePanes('topRow')">Freeze the top row</button>
  <button @click="sheet?.merge('mergeCenter')">Merge and center</button>
  <button @click="sheet?.formatCells()">Format Cells…</button>
  <SheetView ref="sheet" :source="source" :locale="source.locale" style="height: 360px" />
</template>
```

## `SheetToolbar`

Excel's formatting toolbar over the same source; shown only when the source has `format`. See [toolbar.md](toolbar.md).

| Prop | Type | Meaning |
| --- | --- | --- |
| `source` | `ToolbarSource` (the same as `SheetSource`) | The sheet's source. |
| `state` | `SheetViewState \| null` | The payload of the sheet's `state` event; with `null` the buttons are disabled. |
| `messages` | `Partial<SheetMessages>` | Interface texts. |
| `locale` | `string` | The workbook's locale, in which the Number Format box names its categories and items; en-US without it. |
| `sheet` | `SheetViewApi \| null` | The sheet's template ref. With it the toolbar formats, undoes, redoes and merges through the sheet; without it, through the source, and it emits `select` and `error`. |

| Event | Payload | When |
| --- | --- | --- |
| `select` | `{ sheet: string; range: string }` | Where the selection goes after an undo or a redo, for the sheet's `focus`; not emitted with `sheet`. |
| `error` | `string` | The source rejected a call or the call failed; not emitted with `sheet`. |
| `done` | none | A call has completed: the host may return focus to the grid. |
| `formatCells` | none | Format Cells… was clicked. With a listener, the button opens the host's dialog; without one, it opens the sheet's dialog (through `sheet`), and without either there is no button. |

## `FormatDialog`

Excel's Format Cells dialog for the selection, for a host that places its own dialog. `SheetView` has a built-in one; see [toolbar.md](toolbar.md#formatdialog).

| Prop | Type | Meaning |
| --- | --- | --- |
| `open` | `boolean` | Shown while true; each time it opens, it starts from the active cell's formatting. |
| `state` | `SheetViewState \| null` | The payload of the sheet's `state` event; with `null` nothing is shown. |
| `messages` | `Partial<SheetMessages>` | Interface texts. |
| `locale` | `string` | The workbook's locale for the Number tab. |

| Event | Payload | When |
| --- | --- | --- |
| `apply` | `SheetFormatPatch[]` | OK was clicked with changes: patches for the sheet's `format`. |
| `close` | none | OK or Cancel was clicked: the host hides the dialog. |

## The source: `SheetSource`

The component accesses its data only through this interface. Rows and columns are 1-based, cells are A1 addresses, and formulas are displayed and typed in the source's own locale. Two methods are required; each optional method enables a feature. An editing method returns an `EditOutcome`.

A source over a workbook on the host's server: sheets and windows via GET, edits via POST, and changes made elsewhere via server-sent events. The server responds with the contract's types as JSON; everything else (undo, Find, formatting) is added as further optional methods in the same way.

```ts
import type { EditOutcome, SheetInfo, SheetPending, SheetSource, SheetWindow, SourceEvent } from '@xlcalc/vue'

export function serverSource(base: string): SheetSource {
  const get = async <T>(path: string, signal?: AbortSignal): Promise<T> => {
    const res = await fetch(base + path, { signal })
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`)
    return (await res.json()) as T
  }
  const post = async <T>(path: string, body: unknown): Promise<T> => {
    const res = await fetch(base + path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`)
    return (await res.json()) as T
  }
  return {
    sheets: (signal) => get<SheetInfo[]>('/sheets', signal),
    // No visibleRows, so a request never carries a filter
    window: ({ sheet, row, rows, col, cols }, signal) =>
      get<SheetWindow | SheetPending>(`/window?${new URLSearchParams({ sheet, row: `${row}`, rows: `${rows}`, col: `${col}`, cols: `${cols}` })}`, signal),
    edit: (entries, opts) => post<EditOutcome>('/edit', { entries, ...opts }),
    subscribe(listener) {
      const events = new EventSource(`${base}/events`)
      events.onmessage = (e) => listener(JSON.parse(e.data) as SourceEvent)
      return () => events.close()
    },
  }
}
```

### Required

| Method | Meaning |
| --- | --- |
| `sheets(signal?) => Promise<SheetInfo[]>` | The workbook's sheets in order; read again after a structure change and on `reload`. |
| `window(req: WindowRequest, signal?) => Promise<SheetWindow \| SheetPending>` | A rectangular block of a sheet's cells (a window), at most `maxCells` cells. `SheetPending` makes the sheet wait and request it again. |

### Optional methods

| Method | Turns on | Without it | More |
| --- | --- | --- | --- |
| `edit(entries: CellEntry[], opts?: EditOptions)` | Editing. `text: null` clears. | The sheet is read-only. | [editing.md](editing.md) |
| `paste(sheet, cell, text)` | A paste parsed by the source, as a single undo step. | The fields go to `edit`. | [editing.md](editing.md) |
| `fill(sheet, range, text, active, opts?)` | Ctrl+Enter over a selection. | Only the active cell gets the text. | [editing.md](editing.md) |
| `undo()`, `redo()` | Undo and redo. | No undo. | [editing.md](editing.md) |
| `structure(op: SheetOp)` | Structure operations on sheets, rows and columns (insert, delete, rename, move, hide); copying and moving ranges; merging. | No structure menu, no fill handle, and a copy is pasted as text. | [sheets.md](sheets.md), [clipboard.md](clipboard.md) |
| `layout(change: LayoutChange)` | Column widths and row heights, hiding and fitting rows and columns. | A changed width lasts until the page reloads. | [sheets.md](sheets.md) |
| `copyText(sheet, range, opts?) => Promise<string>` | Excel's clipboard text for a range. | The sheet copies the displayed text of its windows. | [clipboard.md](clipboard.md) |
| `format(sheet, patches: FormatPatch[])` | Formatting, the toolbar, Format Cells, Ctrl+B/I/U. | No toolbar. | [toolbar.md](toolbar.md) |
| `subscribe(listener: (e: SourceEvent) => void) => () => void` | Changes made elsewhere. Returns the unsubscribe function. | The sheet sees only its own edits. | [editing.md](editing.md) |
| `validation(sheet, cell, signal?) => Promise<SheetValidation \| null>` | Excel's validation dropdown and input message. | None. | [validation.md](validation.md) |
| `pick(sheet, cell, item, text)` | A picked list item written as Excel writes it. | The item goes to `edit` with `accept`. | [validation.md](validation.md) |
| `freeze(sheet, panes: { rows; cols } \| null)` | Freeze Panes in the context menu. | Only the formula bar's lock button hides the workbook's frozen panes. | [view.md](view.md) |
| `functionNames(signal?) => Promise<string[]>` | Formula AutoComplete and Insert Function. | No list, and the fx button is disabled. | [autocomplete.md](autocomplete.md) |
| `functionSignature(name, signal?) => Promise<SheetFunctionSignature \| null>` | A function's arguments for Excel's argument tooltip while a formula is typed (`SUM(number1, [number2], ...)`), and for Insert Function its syntax, `description` and `category` with `categoryName` (xlcalc 0.42); the name as typed, in any letter case; null for a function the source does not know. | No tooltip; Insert Function lists only the names. | [autocomplete.md](autocomplete.md) |
| `explain(sheet, cell) => Promise<string>` | Trace Precedents. | No panel. | [view.md](view.md) |
| `end(sheet, cell, direction, signal?) => Promise<string>` | Ctrl+arrow as Excel's Range.End. | The edge of the data known to the sheet. | [editing.md](editing.md) |
| `stats(sheet, range, filter?, signal?) => Promise<RangeStats>` | Selection totals over the entire range. | Counted over the loaded windows. | |
| `find(req: FindRequest, signal?) => Promise<FindResult>` | Find. | No search. | [find-replace.md](find-replace.md) |
| `replace(req: FindRequest, by) => Promise<ReplaceOutcome>` | Replace. | No Replace row. | [find-replace.md](find-replace.md) |
| `outline(change: OutlineChange)` | The outline buttons, Group and Ungroup. | The bars are drawn, but their buttons do nothing. | [sheets.md](sheets.md) |
| `visibleRows(sheet, filter, signal?) => Promise<{ rows: number[] \| null }>` | The display filter and sort. | No filter. | [view.md](view.md) |
| `columnValues(sheet, col, filter, signal?) => Promise<{ values: { text; count }[]; more }>` | The filter funnel's list of values. | No list in the funnel. | |

| Property | Type | Meaning |
| --- | --- | --- |
| `maxCells` | `number` | The maximum number of cells in one window, 20,000 by default. |
| `history` | `{ undo: number; redo: number }` | The number of steps `undo` and `redo` can take, for the toolbar buttons. |

### Windows and sheets

A window is a rectangular block of cells that the sheet reads from the source as it scrolls. `WindowRequest` is `{ sheet, row, rows, col, cols, filter? }`; under a display filter, `row` is a position among the visible rows.

`SheetWindow`:

| Field | Type | Meaning |
| --- | --- | --- |
| `row`, `col` | `number` | The window's top-left cell. |
| `text` | `string[][]` | Displayed text, rows × columns; an empty cell is `''`. |
| `values` | `CellValue[][]` | Values of the same cells, for totals and alignment. |
| `formulas` | `Record<string, string>` | Formulas by address, as displayed, without `=`. |
| `editText` | `Record<string, string>?` | The text each cell opens with for editing; a cell missing from it opens empty. Without it the component produces the text from the formula, otherwise from the value's `entryText` (`xlcalc/entry`) according to `numberFormats`. |
| `colors` | `Record<string, string>` | The color from the number format's section, `#RRGGBB`. |
| `styleIds`, `styles` | `Record<string, number>`, `CellStyle[]` | Each cell's style by index into `styles`. |
| `numberFormats` | `Record<string, string>?` | The format code each cell is displayed with; a cell missing from it is General. |
| `widths` | `(number \| null)[]` | Column widths in xlcalc's units, `null` for the default. |
| `rowHeights` | `Record<number, number>` | Row heights in points. |
| `defaultRowHeight` | `number?` | In points. |
| `merges` | `string[]` | Merged ranges. |
| `tables` | `{ name; range; headerRow }[]?` | Tables, whose header and border are drawn. |
| `appearance` | `Record<string, ConditionalAppearance>?` | The cells' conditional formatting, as xlcalc calculates it. |
| `links` | `Record<string, { href?; location?; tooltip?; target? }>?` | Hyperlinks. |
| `notes` | `Record<string, { author?; text }>?` | Notes. |
| `notice` | `string?` | A notice above the grid. |
| `drill` | `Record<string, true>?` | List mode: drill-down cells, underlined; a click or Enter emits `cellDrill`. An empty cell is never a drill-down. |

`SheetPending` is `{ pending: { elapsedMs } }`: the workbook is recalculating, and the sheet requests the window again.

`SheetInfo`:

| Field | Type | Meaning |
| --- | --- | --- |
| `name` | `string` | The sheet's name. |
| `maxRow`, `maxCol` | `number` | The extent of the data. |
| `freeze` | `{ rows; cols }?` | The workbook's frozen panes. |
| `hidden` | `boolean \| 'very'?` | Not shown on the tabs; `'very'` is also not offered for Unhide. |
| `locked` | `{ range; reason }[]?` | Ranges that cannot be edited, with the reason displayed. |
| `columns` | `{ title; align?; sortable? }[]?` | List mode: titles in place of column letters, aligned by `align` (the totals too); clicking a `sortable` title emits `sort`. |
| `rowNumbers` | `boolean?` | Row numbers on the left, shown by default. |
| `totals` | `(string \| null)[]?` | List mode: a totals row pinned to the bottom of the viewport, the displayed text for each of the sheet's columns. |
| `outline` | `SheetOutline?` | Excel's outline: `rows` and `columns` as `OutlineGroup[]` (`{ from, to, level, collapsed }`), `summaryBelow`, `summaryRight`. |

### Edits and outcomes

`CellEntry` is `{ sheet, cell, text: string | null }`. `EditOptions` is `{ accept?: true; inPlace?: true }`: `accept` writes a value that a validation rule rejected after the user chose to keep it, and `inPlace` indicates that the text was edited starting from the cell's own text (F2, a double click, the formula bar), so Excel's automatic percentage entry adds no `%`.

`EditOutcome` is one of:

- `{ ok: true; changes: CellUpdate[]; select?: { sheet; range } }`: `select` is where the selection goes after `paste`, `fill`, `undo` or `redo`.
- `{ ok: false; reason; message; at?; correction?; alert? }`, where `reason` is `'conflict'` (the workbook changed elsewhere), `'refused'` (the cell is locked or the operation is not allowed), `'invalid'` (the formula does not parse; `at` is the position in the typed text, and `correction` is the formula Excel suggests instead in its "We found a typo" dialog, which the sheet offers to accept) or `'failed'` (anything else). `alert` (`SheetValidationAlert`) makes the sheet display Excel's validation alert.

`CellUpdate` is `{ sheet, cell, text, value, color?, formula?, editText?, numberFormat? }`; `formula`, `editText` and `numberFormat` are present only for edited cells.

`SourceEvent`, sent to `subscribe`'s listener:

- `{ kind: 'changed'; changes: CellUpdate[] }`: the loaded windows are patched.
- `{ kind: 'reload'; message? }`: the sheets and windows are read again.
- `{ kind: 'repaint'; ranges: { sheet; range }[] }`: the windows that overlap the ranges are re-read in place.

`FormatPatch` is `{ range, style?: StylePatch | null, numberFormat?: NumberFormat | null }`; `style: null` removes the cells' own style. `LayoutChange` is `{ sheet, columns, width?, hidden?, fit? }` or `{ sheet, rows, height?, hidden?, fit? }`. `SheetOp` is xlcalc's `StructureOp` without its table operations. `OutlineChange` is `{ op: 'group' | 'ungroup', sheet, rows | columns }`, `{ op: 'collapse' | 'expand', sheet, rows | columns, level }` or `{ op: 'showLevel', sheet, axis, level }`.

### Find, filter, validation

| Type | Shape |
| --- | --- |
| `FindRequest` | `{ sheet, query, range?, matchCase?, wholeCell?, lookIn?: 'formulas' \| 'values', byColumns?, after?, filter? }`. `query` accepts Excel's wildcards `*`, `?` and `~`. |
| `FindResult` | `{ hits: { row, col }[]; total }` |
| `ReplaceOutcome` | `EditOutcome & { replaced: number; refused?: { sheet, cell, message }[] }` |
| `SheetFilter` | `{ head: number; cols?: Record<number, { values?: string[]; text?: string }>; sort?: { col; dir: 'asc' \| 'desc' } }` |
| `RangeStats` | `{ cells, numbers, sum, min, max }` |
| `SheetFunctionSignature` | xlcalc's `FunctionSignature` with `description?`, a one-line description of the function in the source's locale, for Insert Function. |
| `SheetValidation` | `{ list?: string[]; prompt?: { title?; text } }` |
| `SheetValidationAlert` | `{ severity: 'stop' \| 'warning' \| 'information'; title?; message; sheet; cell }` |

## Interface texts: `SheetMessages`

`enMessages` (the default) and `ruMessages` are complete sets. A host passes a set, or part of one, as `messages`; the keys and their arguments are declared in `src/contract/messages.ts`.

```ts
import { ruMessages } from '@xlcalc/vue'
const messages = { ...ruMessages, loading: 'Загружаю…' }
```

## `memorySource`

`memorySource(book: MemoryBook, options?: MemorySourceOptions): MemorySource` is a source over an in-memory workbook, for an example or a test. It has no engine: a formula's value is whatever the cell's `value` specifies. Find and Replace search a cell's formula, otherwise its value's `entryText` in en-US (`xlcalc/entry`, loaded on the first search).

| Type | Shape |
| --- | --- |
| `MemoryBook` | `{ sheets: MemorySheet[]; styles?: CellStyle[]; functions?: string[] }`; `functions` enables `functionNames`. |
| `MemorySheet` | `SheetInfo` without `maxRow`, `maxCol` and `outline`, plus `cells: Record<string, MemoryCell \| CellValue>`, `widths?` (keyed by `B`, `B:D` or `C:C`, in xlsx units, 0 hides), `heights?` (keyed by `5`, `5:9` or `5:5`, in points, 0 hides), `defaultRowHeight?`, `merges?`, `tables?`, `notice?`, `outline?: MemoryOutline`, `validations?: MemoryValidation[]`. |
| `MemoryOutline` | `{ rows?: Record<string, number>; columns?: Record<string, number>; collapsed?: string[]; summaryBelow?: boolean; summaryRight?: boolean }`: outline levels (1 to 7) by row, column or span, keyed as in `widths` and `heights` (`'3:9': 1`, `'C:E': 2`), the summary rows and columns marked collapsed (`'10'`, `'F'`), and where the summaries are (below and to the right by default). |
| `MemoryCell` | `{ value?, text?, formula?, style?, color?, numberFormat?, link?, note?, appearance?, drill? }`; `numberFormat` is the code the cell is displayed with (`0.0%`, `dd.mm.yyyy`), which the windows return as `numberFormats`, so a percentage or a date opens for editing as in Excel. |
| `MemoryValidation` | `{ range; list?: string[]; whole?: { min?; max? }; prompt?: { title?; text }; alert?: { severity?; title?; message? } }` |
| `MemorySourceOptions` | `{ delay?: number; maxCells?: number }`; `delay` makes every response take that many milliseconds. |

`MemorySource` is a `SheetSource` with four extra methods for tests and examples: `failNext(...outcomes)` makes the next edits fail with these outcomes, `emit(event)` sends an event to the subscribers, `setCells(sheet, entries)` changes cells as another tab would, and `recompute(ms)` makes windows return `pending` for that long.

A workbook with styles, a formula, a note, a frozen header row, a list validation rule and the functions for Formula AutoComplete:

```ts
import { memorySource } from '@xlcalc/vue'

const source = memorySource(
  {
    styles: [{ font: { bold: true }, fill: { color: '#E8EEF8' } }, { alignment: { horizontal: 'right' } }],
    functions: ['SUM', 'AVERAGE', 'COUNT', 'IF'],
    sheets: [
      {
        name: 'Orders',
        freeze: { rows: 1, cols: 0 },
        widths: { A: 18, B: 12 },
        cells: {
          A1: { value: 'Client', style: 0 },
          B1: { value: 'Amount', style: 0 },
          C1: { value: 'Status', style: 0 },
          A2: 'Acme',
          B2: { value: 1250.5, text: '1,250.50', style: 1 },
          C2: { value: 'Open', note: { author: 'Ann', text: 'Due on Friday' } },
          A3: 'Total',
          B3: { value: 1250.5, text: '1,250.50', formula: 'SUM(B2:B2)', style: 1 },
        },
        validations: [{ range: 'C2:C100', list: ['Open', 'Paid', 'Cancelled'] }],
      },
    ],
  },
  { delay: 100 }, // every response takes 100 ms, to show the sheet waiting
)
```

In a test the extra methods simulate what a server would do:

```ts
import { memorySource } from '@xlcalc/vue'

const source = memorySource({ sheets: [{ name: 'Sheet1', cells: { A1: 1 } }] })
source.failNext({ ok: false, reason: 'conflict', message: 'The workbook changed elsewhere' })
console.log(await source.edit?.([{ sheet: 'Sheet1', cell: 'A1', text: '2' }])) // the conflict
console.log(await source.edit?.([{ sheet: 'Sheet1', cell: 'A1', text: '2' }])) // ok, with A1's change
source.setCells('Sheet1', { B1: '5' }) // subscribers receive it as `changed`
```

## `@xlcalc/vue/local`

A source over an xlcalc workbook calculated in the browser, plus reading and writing xlsx files. See [local.md](local.md).

### `localSource(input: BookInput, opts?: LocalSourceOptions): LocalSource`

The options, `LocalSourceOptions`:

| Option | Type | Meaning |
| --- | --- | --- |
| `calc` | `'pool' \| 'inline' \| Calc` | Where the workbook is calculated: in Web Workers (the default where `Worker` exists), on the page's main thread in slices (the default elsewhere), or by the host's calculator, which the source does not close. |
| `pool` | `CalcPoolOptions` | Options for the pool the source starts, such as `spawn` for a bundler that relocates the thread script. |
| `session` | `AsyncSessionOptions` | Options for the workbook's session; `{ cached }` from `openXlsx` keeps a file's stored values. |
| `maxCells` | `number` | The maximum number of cells in one window, 20,000 by default. |
| `notice` | `string` | A notice included in every window. |

```ts
import { localSource } from '@xlcalc/vue/local'

const source = localSource(
  { sheets: [{ name: 'Sheet1', cells: { A1: { value: 2 }, A2: { formula: 'A1*21' } } }] },
  { calc: 'inline', notice: 'A copy: changes are not saved' }, // calculated on the page's main thread, with a notice above the grid
)
await source.ready // the workbook is calculated
console.log(source.locale, await source.book())
await source.close() // when the sheet is removed, as in onUnmounted
```

`LocalSource` has every method of `SheetSource` (the type marks `copyText` as optional, but the source always has it), plus:

| Member | Type | Meaning |
| --- | --- | --- |
| `locale` | `string` | The workbook's locale tag; pass it to `SheetView`'s `locale`. |
| `ready` | `Promise<void>` | Resolves when the workbook is calculated; rejects if it could not be opened. |
| `history` | `SessionHistory` | The steps `undo` and `redo` can take now. |
| `book()` | `Promise<BookInput>` | The workbook with all changes so far. |
| `snapshot()` | `Promise<BookSnapshot>` | The workbook and its values captured at one moment, as `writeXlsx` takes them. |
| `close()` | `Promise<void>` | Closes the session, and the pool when the source started it. |

### Files

| Function | Returns | Meaning |
| --- | --- | --- |
| `openXlsx(data: Blob \| ArrayBuffer \| Uint8Array, opts?: XlsxIssuesOptions)` | `Promise<XlsxRead>` (`{ book, cached, issues }`) | Reads a file into a workbook and the values stored in it. |
| `saveXlsx(source: LocalSource, opts?: XlsxIssuesOptions & { mode?: 'formulas' \| 'values' })` | `Promise<XlsxWritten>` (`{ data, issues }`) | Writes the source's workbook with its values; `values` mode writes no formulas. |
| `groupIssues(issues: XlsxIssue[])` | `IssueGroup[]` | Issues grouped by kind, in order of first appearance, for a notice of what a file lost. |

`XlsxIssuesOptions` is `{ locale?: string | Locale }`: the language of the issue messages, English by default. `IssueGroup` is `{ code, count, sheets: string[], message }`, where `message` is the first issue's message.

## CSS variables

`@xlcalc/vue/style.css` styles `.xs-sheet`, `.xs-toolbar` and the dialog with `--xs-*` variables, declared on `.xs-sheet`, `.xs-toolbar` and `.xs-fmt-back` (a `FormatDialog` placed by the host) themselves. A host restyles them, or switches them to a dark theme, by overriding the variables on those elements with a selector more specific than the package's, such as one scoped under its own class; a value set only on an ancestor does not reach them.

```css
/* A dark sheet, toolbar and Format Cells under the host's .dark */
.dark .xs-sheet,
.dark .xs-toolbar,
.dark .xs-fmt-back {
  --xs-bg: #1e1f22;
  --xs-fg: #e3e3e3;
  --xs-muted: #9aa0a6;
  --xs-grid: #3a3b3e;
  --xs-head-bg: #2b2c2f;
  --xs-head-fg: #b0b3b8;
  --xs-head-on-bg: #2d4a75;
  --xs-head-on-fg: #d3e3fd;
  --xs-bar-bg: #26272a;
  --xs-link: #8ab4f8;
  --xs-font-size: 13px;
}
```

| Variable | Default | What |
| --- | --- | --- |
| `--xs-bg`, `--xs-fg`, `--xs-muted` | `#ffffff`, `#1f2328`, `#6b7280` | Background, text, secondary text. |
| `--xs-font`, `--xs-font-size` | `system-ui, …`, `12.5px` | The interface font. |
| `--xs-grid` | `#e1e3e6` | Gridlines and interface borders. |
| `--xs-head-bg`, `--xs-head-fg` | `#f3f4f6`, `#5f6368` | Row numbers and column letters. |
| `--xs-head-on-bg`, `--xs-head-on-fg` | `#d3e3fd`, `#0b57d0` | Headers of selected rows and columns. |
| `--xs-accent`, `--xs-sel-tint` | `#1a73e8`, `rgb(26 115 232 / 0.12)` | The selection border and tint. |
| `--xs-on-accent` | `#ffffff` | Text on the accent color: a primary button, a list's highlighted item. |
| `--xs-link`, `--xs-note` | `#1558d6`, `#dc2626` | Hyperlinks, a note indicator. |
| `--xs-flash`, `--xs-hit` | `rgb(250 204 21 / 0.85)`, `rgb(250 204 21 / 0.45)` | The flash after a jump, Find matches. |
| `--xs-freeze-line`, `--xs-border-auto` | `#9aa0a6`, `#000000` | The frozen panes line, a border with automatic color. |
| `--xs-table-head-bg`, `--xs-table-line` | `#e8eef8`, `#8faadc` | Tables. |
| `--xs-locked`, `--xs-pending`, `--xs-error` | `#8b8f96`, `#8b8f96`, `#b42318` | Locked cells, pending edits, errors. |
| `--xs-bar-bg` | `#fafafa` | The formula bar, the toolbar and the Trace Precedents panel. |
| `--xs-message-bg`, `--xs-message-fg` | `#1f2328`, `#ffffff` | The message line. |
| `--xs-copy-line`, `--xs-fill-frame` | `#1a73e8`, `#5f6368` | The moving border of a copy, the fill handle's outline. |
| `--xs-shadow-pop`, `--xs-shadow-menu`, `--xs-shadow-modal` | `0 2px 6px rgb(0 0 0 / 0.2)`, `0 6px 20px rgb(0 0 0 / 0.18)`, `0 10px 30px rgb(0 0 0 / 0.25)` | Shadows of a list, a tooltip or a box under a cell; of a menu, the filter funnel's list and the message line; of a dialog or an alert. |
| `--xs-backdrop` | `rgb(0 0 0 / 0.25)` | The dimmed page under a dialog or an alert. |
| `--xs-z-menu`, `--xs-z-modal`, `--xs-z-alert` | `40`, `50`, `55` | Stacking layers above the page: a menu or the filter funnel's list, a dialog, an alert (above a dialog). |
| `--xs-outline-bg`, `--xs-outline-line` | `var(--xs-bar-bg)`, `#8c8c8c` | The outline bars and brackets. |
| `--xs-outline-button`, `--xs-outline-tick` | `13px`, `5px` | The outline button size and the bracket tick length. |
| `--xs-outline-button-bg`, `--xs-outline-button-fg`, `--xs-outline-button-border` | `var(--xs-bg)`, `#3c4043`, `#9aa0a6` | The outline buttons. |
| `--xs-explain-w`, `--xs-explain-h` | `340px`, `45%` | The Trace Precedents panel: its width, and its height on a phone. |
| `--xs-prompt-bg`, `--xs-prompt-fg`, `--xs-prompt-line` | `#ffffe1`, `#1f2328`, `#767676` | The data validation input message. |
| `--xs-alert-stop`, `--xs-alert-warning`, `--xs-alert-information` | `#d13438`, `#ffb900`, `#0f6cbd` | The validation alert's icons. |
| `--xs-fl-line` | `#7a7a7a` | The border of lists under a cell: Formula AutoComplete and validation lists. |
