# The formatting toolbar and Format Cells

## `SheetToolbar`

`SheetToolbar` is Excel's formatting row over the same source: undo and redo, bold, italic, underline, font and fill colours from Excel's palette, borders (all, outside, bottom, none), horizontal alignment, wrap text, Excel's Number Format box and, for a host that shows the dialog ([below](#formatdialog)), Format Cells. It formats the selection the sheet's `state` event gives, through `format(sheet, patches)`, as Excel formats a selection: bold goes over each cell's look, outside borders are four patches in one call and so one step of undo. Its buttons are pressed by what the active cell shows, and pressed again they take it off, as Excel's are. It shows only when the source has `format`; undo and redo come with the source's `undo` and `redo`, off while its `history` (the local source's) has no step.

With `sheet` and the source's `structure` the toolbar has Excel's Merge & Center: the button merges and centres the selection, and its list has Merge & Center, Merge Across, Merge Cells and Unmerge Cells, each one step of undo through the sheet's `merge(command)`. The button is pressed on a merged cell, and pressed then it unmerges, as Excel's; on one cell outside a merged area nothing happens, and it is off on a read-only sheet and under the display filter. Before a merge that would keep one value of several in an area (in a row, for Merge Across) the sheet shows Excel's warning, "Merging cells only keeps the upper-left value and discards other values.", OK on Enter and Cancel on Esc; it counts the cells holding a value or a formula in its loaded windows, or in the source's windows over the range when they do not hold it. A refusal shows the source's message.

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

// A source with format, undo and structure: the toolbar shows only over a source with format
const source = localSource({
  sheets: [{ name: 'Sheet1', cells: { A1: { value: 'Region' }, B1: { value: 'Sales' }, A2: { value: 'North' }, B2: { value: 0.125 }, A3: { value: 'South' }, B3: { value: 0.3 } } }],
})
onUnmounted(() => void source.close())

const state = ref<SheetViewState | null>(null)
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <SheetToolbar :source="source" :sheet="sheet" :state="state" :locale="source.locale" />
  <SheetView :source="source" ref="sheet" :locale="source.locale" style="height: 360px" @state="state = $event" />
</template>
```

The Number Format box is Excel's, as xlcalc measured it (`numberFormatCategory` and `numberFormatChoices` of `xlcalc/text`, without the engine): it names the category of the active cell's format (General without one, Custom where no category of Excel's holds the code), and its list applies Excel's eleven items (General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text), each by the format Excel sets (Percentage is `0.00%`), after which the box names the new category. `locale` is the book's locale, as `SheetView`'s: the names and the Currency item's code are its, ru-RU's loaded as a chunk of its own; en-US without it. With `sheet` (the sheet's template ref) the toolbar formats, undoes and redoes through the sheet's methods ([api.md](api.md#sheetviewapi)): the sheet sets the selection and shows a refusal itself. Without it the toolbar calls the source and emits `select` with where the selection goes after an undo or a redo (for the sheet's `focus`) and `error` with a source's refusal. It emits `done` after each call, for the host to give the grid its focus back. Its buttons do not take the focus from the grid. Under 600 px it keeps undo, redo, bold, italic and underline in a row and the rest in a menu; on a narrow frame or under a finger every target is 44 × 44 px.

## `FormatDialog`

`SheetView` has Excel's Format Cells window of its own when the source has `format`: Control+1 opens it (on a Mac too, as the browser keeps Cmd+1), as do the sheet's `formatCells()` and the toolbar's Format Cells… with `sheet`, and OK formats the selection through the sheet's `format`. The window is also exported as `FormatDialog`, for a host that places its own. It is Excel's Format Cells over the selection the sheet's `state` event gives, with the tabs whose settings xlcalc's cell style holds. Number is Excel's Number Format box: the category of the code in the field, the list's items (an item puts its format's code in the field) and the code itself, which may be any. Alignment has the horizontal and vertical alignments, the indent (with left, right and distributed, as in Excel) and wrapping; Font the font's name, Regular, Italic, Bold or Bold Italic, the size, the underline, strikethrough and the colour; Border a line's kind and colour, the presets None, Outline and Inside and a button for each edge round a preview, the inside ones on a selection of more than one row or column; Fill the background colour. Colours are Excel's palette, as on the toolbar, or any `#RRGGBB`. It opens at the active cell's look, with the edges of one cell as its own sides and a range's as not known, as Excel greys them. OK emits `apply` with patches for `SheetView.format` for what changed alone: the font, the alignment, the fill and the number format over the selection, each border edge on its row, column or the cells on both sides of an inside line (one patch over every cell when all edges take one line), all in one call and so one step of undo; then `close`, as Cancel, Esc and a press outside the window do. Tab and Shift+Tab stay inside it. A field past Excel's limits (a size of 1 to 409 points, an indent of 0 to 250) or a colour code that is none holds OK back. On a phone it takes the whole screen with 44 px targets.

A host with a dialog of its own listens to the toolbar's `formatCells`, which the button then emits instead:

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { FormatDialog, SheetToolbar, SheetView, ruMessages, type SheetViewApi, type SheetViewState } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

// A book in ru-RU: its locale module loads as the source opens
const source = localSource({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 1234.5 }, A2: { value: 0.25 } } }] })
onUnmounted(() => void source.close())

const state = ref<SheetViewState | null>(null)
const sheet = ref<SheetViewApi | null>(null)
const formatOpen = ref(false)
</script>

<template>
  <SheetToolbar :source="source" :sheet="sheet" :state="state" :messages="ruMessages" :locale="source.locale" @format-cells="formatOpen = true" />
  <SheetView :source="source" ref="sheet" :messages="ruMessages" :locale="source.locale" style="height: 360px" @state="state = $event" />
  <!-- The host's dialog, in the host's place: OK formats through the sheet -->
  <FormatDialog :open="formatOpen" :state="state" :messages="ruMessages" :locale="source.locale" @apply="sheet?.format($event)" @close="formatOpen = false" />
</template>
```

Patterns of a fill, diagonal borders, rotated and shrunk text, merging, text direction, superscript, subscript and the Protection tab are not there: xlcalc's cell style has none of them.
