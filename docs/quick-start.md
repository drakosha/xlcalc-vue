# Quick start

Common uses of the component, each with a single-file Vue component that works as written. Install the package first ([README](../README.md#install)); every example imports the stylesheet once, which an app does in its entry file instead. Each section links to the full documentation.

## A table the page already has

The host's data shown as a sheet: `memorySource` turns it into a workbook in memory, without the calculation engine. A cell displays its `text` when it has one, so the host formats numbers itself; `readOnly` turns editing off, and the frozen first row keeps the titles in view. The sheet's `state` event reports the selection and the active cell's text, as Excel's Name Box and formula bar show them.

```vue
<script setup lang="ts">
import { shallowRef } from 'vue'
import { SheetView, memorySource, type MemoryCell, type SheetViewState } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const orders = [
  { client: 'Acme', amount: 1250.5, paid: true },
  { client: 'Globex', amount: 980, paid: false },
  { client: 'Initech', amount: 15300.75, paid: true },
]
const money = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

// The titles in row 1 in bold, each order in a row below them
const cells: Record<string, MemoryCell> = { A1: { value: 'Client', style: 0 }, B1: { value: 'Amount', style: 0 }, C1: { value: 'Paid', style: 0 } }
orders.forEach((order, i) => {
  cells[`A${i + 2}`] = { value: order.client }
  cells[`B${i + 2}`] = { value: order.amount, text: money.format(order.amount) }
  cells[`C${i + 2}`] = { value: order.paid }
})

const source = memorySource({
  styles: [{ font: { bold: true } }],
  sheets: [{ name: 'Orders', freeze: { rows: 1, cols: 0 }, widths: { A: 16, B: 12 }, cells }],
})
const state = shallowRef<SheetViewState | null>(null)
</script>

<template>
  <p>{{ state ? `${state.ref}: ${state.value ?? ''}` : '' }}</p>
  <SheetView :source="source" read-only hide-lone-tab style="height: 300px" @state="state = $event" />
</template>
```

Without `readOnly` the cells can be edited, and `memorySource` keeps what is typed; it does not calculate formulas, so a formula cell displays the `value` given with it. [The `memorySource` reference](api.md#memorysource) lists what a sheet can carry: styles, merged ranges, notes, hyperlinks, data validation, outline.

## A model the user edits

`localSource` from `@xlcalc/vue/local` calculates an xlcalc workbook in the browser, in a Web Worker: the user types into the input cells, and the formulas recalculate as in Excel, with undo, fill, Find and the context menus. The workbook is the library's `BookInput`, formulas without `=`, so it can come from the host's code, its server or an xlsx file. This entry point needs `xlcalc` installed.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

// A loan: the inputs in B1:B3, the payment and the interest calculated from them
const source = localSource({
  sheets: [
    {
      name: 'Loan',
      columns: { A: { width: 12 }, B: { width: 14 } },
      cells: {
        A1: { value: 'Amount' }, B1: { value: 200000, numberFormat: '#,##0' },
        A2: { value: 'Rate' }, B2: { value: 0.06, numberFormat: '0.00%' },
        A3: { value: 'Years' }, B3: { value: 20 },
        A5: { value: 'Payment' }, B5: { formula: '-PMT(B2/12,B3*12,B1)', numberFormat: '#,##0.00' },
        A6: { value: 'Interest' }, B6: { formula: 'B5*B3*12-B1', numberFormat: '#,##0.00' },
      },
    },
  ],
})
// The worker and the session end with the page's component
onUnmounted(() => void source.close())
</script>

<template>
  <SheetView :source="source" :locale="source.locale" hide-lone-tab style="height: 300px" />
</template>
```

`source.book()` returns the workbook with the user's changes, for the host to store. [In-browser calculation](local.md) describes what the local source does and its options, such as `calc: 'inline'` to calculate on the page's main thread.

## Excel files: open, edit, save

`openXlsx` reads a file into a workbook and the values Excel stored in it, and `saveXlsx` writes the workbook back, formulas along with their values. A new source replaces the previous workbook; the previous source is closed.

```vue
<script setup lang="ts">
import { onUnmounted, shallowRef } from 'vue'
import { SheetView } from '@xlcalc/vue'
import { localSource, openXlsx, saveXlsx, type LocalSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = shallowRef<LocalSource | null>(null)

async function open(event: Event): Promise<void> {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (!file) return
  const { book, cached } = await openXlsx(file)
  void source.value?.close()
  source.value = localSource(book, { session: { cached } })
}

async function save(): Promise<void> {
  if (!source.value) return
  const { data } = await saveXlsx(source.value)
  const link = document.createElement('a')
  link.href = URL.createObjectURL(new Blob([data as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
  link.download = 'workbook.xlsx'
  link.click()
  setTimeout(() => URL.revokeObjectURL(link.href))
}

onUnmounted(() => void source.value?.close())
</script>

<template>
  <input type="file" accept=".xlsx" @change="open" />
  <button :disabled="!source" @click="save">Save as</button>
  <SheetView v-if="source" :source="source" :locale="source.locale" style="height: 420px" />
</template>
```

`openXlsx` also returns `issues`, what the workbook could not keep from the file (pictures, charts); [Files](local.md#files) shows how to report them, and the [example app](example-app.md) is a complete Excel-like page built this way.

## Excel's formatting toolbar

`SheetToolbar` is Excel's formatting toolbar for the sheet's selection: Bold, colors, borders, alignment, Wrap Text, the Number Format box, Merge & Center, Undo and Redo. It reads the selection from the sheet's `state` event and formats through the sheet's template ref. It is shown for a source with `format`, such as the local one; Ctrl+1 opens Excel's Format Cells dialog on the sheet itself.

```vue
<script setup lang="ts">
import { onUnmounted, ref, shallowRef } from 'vue'
import { SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({
  sheets: [{ name: 'Sheet1', cells: { A1: { value: 'Region' }, B1: { value: 'Sales' }, A2: { value: 'North' }, B2: { value: 1200 }, A3: { value: 'South' }, B3: { value: 950 } } }],
})
onUnmounted(() => void source.close())

const sheet = ref<SheetViewApi | null>(null)
const state = shallowRef<SheetViewState | null>(null)
</script>

<template>
  <SheetToolbar :source="source" :sheet="sheet" :state="state" :locale="source.locale" />
  <SheetView ref="sheet" :source="source" :locale="source.locale" style="height: 360px" @state="state = $event" />
</template>
```

[The formatting toolbar and Format Cells](toolbar.md) describes each button and how a host places its own Format Cells dialog.

## A workbook on your server

The component reads a workbook only through a `SheetSource`, so a workbook that lives on the host's server needs two methods: the list of sheets and a window of cells. Each further method turns on a feature: `edit` makes the sheet editable, `subscribe` brings in changes made elsewhere, `undo`, `find`, `format` and the rest add their commands. The `fallback` slot shows a failure to read the workbook, with a retry.

```vue
<script setup lang="ts">
import { SheetView, type SheetInfo, type SheetPending, type SheetSource, type SheetWindow } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const props = defineProps<{ base: string }>()

async function get<T>(path: string, signal?: AbortSignal): Promise<T> {
  const res = await fetch(props.base + path, { signal })
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`)
  return (await res.json()) as T
}

// The server answers with the contract's types as JSON; without `edit` the sheet is read-only
const source: SheetSource = {
  sheets: (signal) => get<SheetInfo[]>('/sheets', signal),
  window: ({ sheet, row, rows, col, cols }, signal) =>
    get<SheetWindow | SheetPending>(`/window?${new URLSearchParams({ sheet, row: `${row}`, rows: `${rows}`, col: `${col}`, cols: `${cols}` })}`, signal),
}
</script>

<template>
  <SheetView :source="source" style="height: 420px">
    <template #fallback="{ failure, retry }">
      <p>{{ failure.text }} <button @click="retry">Try again</button></p>
    </template>
  </SheetView>
</template>
```

The server can calculate with xlcalc itself: a session's `getViewport` returns what a window holds. [The source](api.md#the-source-sheetsource) lists every method with the feature it turns on, with a fuller source over GET, POST and server-sent events; a new `source` prop means a new workbook.

## The Russian interface and a Russian workbook

The interface texts are English by default; `ruMessages` is the complete Russian set, and a host overrides single texts by spreading it. A workbook's locale is separate: a ru-RU workbook displays and accepts formulas, numbers and dates as Russian Excel does (`=СУММ(B2:B3)`, `1 234,50`), and the sheet takes it as `locale`.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView, ruMessages } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const messages = { ...ruMessages, loading: 'Загружаю книгу…' }
const source = localSource({
  locale: 'ru-RU',
  sheets: [{ name: 'Лист1', cells: { A1: { value: 'Выручка' }, B2: { value: 1234.5 }, B3: { value: 800 }, B4: { formula: 'SUM(B2:B3)', numberFormat: '#,##0.00' } } }],
})
onUnmounted(() => void source.close())
</script>

<template>
  <SheetView :source="source" :locale="source.locale" :messages="messages" style="height: 300px" />
</template>
```

[Interface texts](api.md#interface-texts-sheetmessages) lists the keys; the colors and sizes are CSS variables ([CSS variables](api.md#css-variables)).

## Where to go next

- [Cell editing and keyboard shortcuts](editing.md): what each key does and which source method each feature needs.
- [Rows, columns and sheets](sheets.md), [Clipboard and fill handle](clipboard.md), [Find and Replace](find-replace.md), [Data validation](validation.md).
- [List mode](list-mode.md): the host's own list shown on the sheet, sorted by the host, with drill-down cells and a totals row.
- [API reference](api.md): every export, prop, event and method.
