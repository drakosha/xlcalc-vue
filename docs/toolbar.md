# The formatting toolbar and Format Cells

## `SheetToolbar`

`SheetToolbar` is Excel's formatting toolbar over the same source: Undo and Redo, Bold, Italic, Underline, font and fill colors from Excel's palette, borders (All Borders, Outside Borders, Bottom Border, No Border), horizontal alignment, Wrap Text, Excel's Number Format box and, for a host that displays the dialog ([below](#formatdialog)), Format Cells. It formats the selection reported by the sheet's `state` event through `format(sheet, patches)`, as Excel formats a selection: Bold applies on top of each cell's formatting, and Outside Borders is four patches in one call and therefore a single undo step. Its buttons show as pressed according to the active cell's formatting, and clicking a pressed button removes that formatting, as in Excel. The toolbar is shown only when the source has `format`; Undo and Redo appear when the source has `undo` and `redo`, and are disabled while its `history` (the local source's) has no steps.

With `sheet` and the source's `structure`, the toolbar has Excel's Merge & Center: the button merges the selection and centers its contents, and its dropdown has Merge & Center, Merge Across, Merge Cells and Unmerge Cells, each a single undo step through the sheet's `merge(command)`. The button shows as pressed on a merged cell, and clicking it then unmerges, as Excel's does; on a single cell outside a merged range it does nothing, and it is disabled on a read-only sheet and under a display filter. Before a merge that would keep only one of several values in a range (in a row, for Merge Across), the sheet shows Excel's warning "Merging cells only keeps the upper-left value and discards other values.", with OK on Enter and Cancel on Esc; it counts the cells that contain a value or a formula in its loaded windows, or in the source's windows over the range when the loaded ones do not cover it. If the source rejects the merge, its message is displayed.

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

// A source with format, undo and structure: the toolbar is shown only for a source with format
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

The Number Format box works as Excel's does, as measured by xlcalc (`numberFormatCategory` and `numberFormatChoices` of `xlcalc/text`, without the engine): it shows the category of the active cell's format (General if there is none, Custom where no Excel category matches the code), and its list applies Excel's eleven items (General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text), each with the format Excel sets (Percentage is `0.00%`), after which the box shows the new category. `locale` is the workbook's locale, as for `SheetView`: the category names and the Currency item's code follow it, with ru-RU's data loaded as a separate chunk; without it, en-US. With `sheet` (the sheet's template ref), the toolbar formats, undoes and redoes through the sheet's methods ([api.md](api.md#sheetviewapi)), and the sheet sets the selection and displays a rejection itself. Without it, the toolbar calls the source and emits `select` with where the selection goes after an undo or a redo (for the sheet's `focus`) and `error` with the source's rejection. It emits `done` after each call, so the host can return focus to the grid. Its buttons do not take focus from the grid. Below 600 px it keeps Undo, Redo, Bold, Italic and Underline in a row and moves the rest into a menu; in a narrow viewport or on a touch screen every target is 44 × 44 px.

## `FormatDialog`

`SheetView` has its own Excel Format Cells dialog when the source has `format`: Ctrl+1 opens it (on a Mac as well, since the browser reserves Cmd+1), as do the sheet's `formatCells()` and the toolbar's Format Cells… with `sheet`, and OK formats the selection through the sheet's `format`. The dialog is also exported as `FormatDialog`, for a host that places its own. It is Excel's Format Cells for the selection reported by the sheet's `state` event, with the tabs whose settings xlcalc's cell style supports:

- Number is Excel's Number Format box: the category of the code in the input, the list's items (an item puts its format's code in the input) and the code itself, which may be any code.
- Alignment has the horizontal and vertical alignment, the indent (with Left, Right and Distributed alignment, as in Excel) and text wrapping.
- Font has the font name, Regular, Italic, Bold or Bold Italic, the size, the underline, strikethrough and the color.
- Border has the line style and color, the presets None, Outline and Inside, and a button for each edge around a preview, the inside ones for a selection of more than one row or column.
- Fill has the background color.

Colors come from Excel's palette, as on the toolbar, or can be any `#RRGGBB`. The dialog opens with the active cell's formatting; for a single cell the edges show its own borders, and for a range they show as unknown, grayed out as in Excel. OK emits `apply` with patches for `SheetView.format` covering only what changed: the font, the alignment, the fill and the number format over the selection, and each border edge on its row or column, or on the cells on both sides of an inside line (a single patch over every cell when all edges use the same line), all in one call and therefore a single undo step; then it emits `close`, as Cancel, Esc and a click outside the dialog do. Tab and Shift+Tab stay inside it. An input beyond Excel's limits (a size of 1 to 409 points, an indent of 0 to 250) or an invalid color code blocks OK. On a phone the dialog takes the full screen with 44 px targets.

A host with its own dialog listens to the toolbar's `formatCells` event, which the button then emits instead:

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { FormatDialog, SheetToolbar, SheetView, ruMessages, type SheetViewApi, type SheetViewState } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

// A workbook in ru-RU: its locale module loads when the source opens
const source = localSource({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 1234.5 }, A2: { value: 0.25 } } }] })
onUnmounted(() => void source.close())

const state = ref<SheetViewState | null>(null)
const sheet = ref<SheetViewApi | null>(null)
const formatOpen = ref(false)
</script>

<template>
  <SheetToolbar :source="source" :sheet="sheet" :state="state" :messages="ruMessages" :locale="source.locale" @format-cells="formatOpen = true" />
  <SheetView :source="source" ref="sheet" :messages="ruMessages" :locale="source.locale" style="height: 360px" @state="state = $event" />
  <!-- The host's own dialog, placed by the host: OK formats through the sheet -->
  <FormatDialog :open="formatOpen" :state="state" :messages="ruMessages" :locale="source.locale" @apply="sheet?.format($event)" @close="formatOpen = false" />
</template>
```

Fill patterns, diagonal borders, rotated text, shrink to fit, merging, text direction, superscript, subscript and the Protection tab are not available, since xlcalc's cell style does not support them.
