# The formatting toolbar and Format Cells

## `SheetToolbar`

`SheetToolbar` is Excel's formatting toolbar over the same source: Undo and Redo, the Font and Font Size boxes with Increase and Decrease Font Size, Bold, Italic, Underline, font and fill colors from Excel's palette, borders (All Borders, Outside Borders, Bottom Border, No Border), Top, Middle and Bottom Align, horizontal alignment, Decrease and Increase Indent, Wrap Text, Excel's Number Format box with Accounting Number Format, Percent Style, Comma Style, Increase Decimal and Decrease Decimal, for a host that displays the dialog ([below](#formatdialog)) Format Cells, and with `sheet` Excel's AutoSum. It formats the selection reported by the sheet's `state` event through `format(sheet, patches)`, as Excel formats a selection: Bold applies on top of each cell's formatting, and Outside Borders is four patches in one call and therefore a single undo step. Its buttons show as pressed according to the active cell's formatting, and clicking a pressed button removes that formatting, as in Excel. The toolbar is shown only when the source has `format`; Undo and Redo appear when the source has `undo` and `redo`, and are disabled while its `history` (the local source's) has no steps.

With `sheet` and the source's `structure`, the toolbar has Excel's Merge & Center: the button merges the selection and centers its contents, and its dropdown has Merge & Center, Merge Across, Merge Cells and Unmerge Cells, each a single undo step through the sheet's `merge(command)`. The button shows as pressed on a merged cell, and clicking it then unmerges, as Excel's does; on a single cell outside a merged range it does nothing, and it is disabled on a read-only sheet and under a display filter. Before a merge that would keep only one of several values in a range (in a row, for Merge Across), the sheet shows Excel's warning "Merging cells only keeps the upper-left value and discards other values.", with OK on Enter and Cancel on Esc; it counts the cells that contain a value or a formula in its loaded windows, or in the source's windows over the range when the loaded ones do not cover it. If the source rejects the merge, its message is displayed.

With `sheet`, the toolbar has Excel's AutoSum from Home > Editing, a split button like Merge & Center: Σ enters Sum in the active cell through the sheet's `autoSum()`, opening its edit with `=SUM(range)` and the range selected ([editing.md](editing.md#autosum)), and its dropdown lists Sum, then Average, Count Numbers, Max and Min as the source lists them in `autoSumFunctions` (`./local` lists all four), and More Functions…, which opens the sheet's Insert Function dialog (`insertFunction()`) for a source with `functionNames`. Both parts are disabled while `state.canAutoSum` is absent: without the source's `autoSum` and on a read-only sheet. Its texts are `autoSum`, `autoSumMenu`, `autoSumSum`, `autoSumAverage`, `autoSumCount`, `autoSumMax`, `autoSumMin` and `autoSumMore`.

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

The Font and Font Size boxes show the active cell's font and size, and each takes a typed value or a pick from its list: Enter or a pick applies it to the selection, and Esc, or leaving the box without Enter, restores what it showed (Esc also emits `done`, so the host can return focus to the grid). A cell without its own font or size shows the workbook's default font (`state.font`, from the source's `SheetInfo.font`), and empty boxes when the source does not give one. A browser does not tell which fonts are installed, so the Font box's list is Excel's common fonts (Aptos, Aptos Narrow, Arial, Calibri, Cambria, Consolas, Courier New, Georgia, Segoe UI, Tahoma, Times New Roman, Verdana), each drawn in its own font, and the `fonts` prop replaces it; any name may be typed, and typing completes it with the first font of the list it begins, as Excel's box does. The Font Size box's list is Excel's (8 to 72), and it takes any size from 1 to 409, with a decimal point or comma, rounded to half points (Excel takes sizes in multiples of .5); any other text stays in the box, marked invalid with the message `fmtSizeRange`, until it is corrected or Esc restores the box. The arrows go through the list, putting each item into the box.

Increase and Decrease Font Size step the active cell's size (11 points for a cell without one, the size the grid draws it at) through Excel's list, beyond 72 in steps of 10 up to 409, and below 8 a point at a time down to 1, and set the new size over the selection; at a limit they send nothing. Top, Middle and Bottom Align show as pressed according to the active cell's vertical alignment, Bottom for a cell without one (Excel's default); clicking a pressed one returns the selection to Bottom. Increase Indent adds a level to the active cell's indent and keeps the Left, Right and Distributed alignments, which have an indent, while any other alignment (General, Center and the rest) becomes Left, as in Excel; Decrease Indent takes a level off without changing the alignment, and does nothing at 0 or with an alignment that has no indent. Indents go up to 250 levels. Percent Style sets Excel's `0%` (the built-in format 9) over the selection, as Ctrl+Shift+% on the grid does ([editing.md](editing.md)); the Number Format box then names Percentage. Accounting Number Format and Comma Style set Excel's built-in formats 44 and 43, which display in the workbook's locale (its currency for Accounting). Increase Decimal and Decrease Decimal take the active cell's format, value and column width (`state.numberFormat`, `cellValue`, `columnWidth`) through xlcalc's `stepDecimals`, as Excel for Mac steps them (a General cell takes the decimals it shows, one more or one less; a code takes a decimal in each section), and put the result on the whole selection, as Excel does; when the step changes nothing (General without a number, no decimal to take away) nothing is sent.

The Number Format box works as Excel's does, as measured by xlcalc (`numberFormatCategory` and `numberFormatChoices` of `xlcalc/text`, without the engine): it shows the category of the active cell's format (General if there is none, Custom where no Excel category matches the code), and its list applies Excel's eleven items (General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text), each with the format Excel sets (Percentage is `0.00%`), after which the box shows the new category. `locale` is the workbook's locale, as for `SheetView`: the category names and the Currency item's code follow it, with ru-RU's data loaded as a separate chunk; without it, en-US. With `sheet` (the sheet's template ref), the toolbar formats, undoes and redoes through the sheet's methods ([api.md](api.md#sheetviewapi)), and the sheet sets the selection and displays a rejection itself. Without it, the toolbar calls the source and emits `select` with where the selection goes after an undo or a redo (for the sheet's `focus`) and `error` with the source's rejection. It emits `done` after each call, so the host can return focus to the grid. Its buttons do not take focus from the grid; the font boxes and the Number Format box do, while they are used. Below 600 px it keeps what fits in a row and moves the rest, the font boxes too, into a menu ([below](#choosing-the-buttons)); in a narrow viewport or on a touch screen every target is 44 × 44 px. Its buttons and boxes have Excel's ScreenTips instead of the browser's `title`: the name and the key (`Bold (Ctrl+B)`, from `SheetMessages`) below the control after the mouse rests on it for 0.4 s, and at once on the next one within 0.6 s, disabled buttons included; a press, a menu opening or the mouse leaving hides it, and touch gets none. The text is each control's `data-tip` (its `aria-label` names it too), and a host's button shows its `label`.

### Choosing the buttons

The `items` prop picks the toolbar's groups and their order, and adds the host's buttons among them. Without it the toolbar draws every group in the default order, `TOOLBAR_GROUPS`: `undo`, `fontStyle` (Bold, Italic, Underline), `font` (the Font and Font Size boxes with Increase and Decrease Font Size), `number` (the Number Format box and its buttons), `colors`, `borders`, `verticalAlign`, `align` (horizontal alignment, indent and Wrap Text), `merge`, `formatCells` and `autoSum`. A group named in `items` still needs what it needs without it (Undo needs the source's `undo`, Merge & Center and AutoSum need `sheet`); an unknown or repeated group is skipped.

A host's button is an object, `{ id, label, icon?, run, disabled?, pressed? }`. The toolbar draws it as its own buttons: `label` is the tooltip and the accessible name, `icon` a Vue component (it gets the class `xs-icon`) or SVG markup, drawn at the toolbar's icon size, and without an icon the label is the button's text. A click closes the toolbar's menus and calls `run`; `disabled` and `pressed` are the host's, so a host that passes a reactive array updates them. Buttons next to each other in `items` form one group. Each has the hook `data-tb-host="<id>"`.

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { SheetToolbar, SheetView, TOOLBAR_GROUPS, type SheetViewApi, type SheetViewState, type ToolbarItem } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({ sheets: [{ name: 'Sheet1', cells: { B2: { value: 120 }, B3: { value: 80 } } }] })
onUnmounted(() => void source.close())

const state = ref<SheetViewState | null>(null)
const sheet = ref<SheetViewApi | null>(null)
// No borders or colors; AutoSum first; the host's Save button at the end, drawn as the toolbar's own
const items: ToolbarItem[] = [
  'autoSum',
  ...TOOLBAR_GROUPS.filter((g) => g !== 'autoSum' && g !== 'borders' && g !== 'colors'),
  { id: 'save', label: 'Save', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 3h11l3 3v15H5z"/><path d="M8 3v6h8V3"/></svg>', run: () => void source.book() },
]
</script>

<template>
  <SheetToolbar :source="source" :sheet="sheet" :state="state" :locale="source.locale" :items="items" />
  <SheetView :source="source" ref="sheet" :locale="source.locale" style="height: 360px" @state="state = $event" />
</template>
```

Below 600 px the toolbar keeps in its row the items that fit, in the order `items` gives them, and moves the rest, the host's buttons included, into the menu under a More button, in order: it measures the row with every item and keeps those from the start that fit beside More, so with the default items a phone keeps Undo, Redo, Bold, Italic and Underline, and a host's few buttons after `undo` stay in the row. When every item fits there is no More button. It measures again when its width, the items or the texts change.

## `FormatDialog`

`SheetView` has its own Excel Format Cells dialog when the source has `format`: Ctrl+1 opens it (on a Mac as well, since the browser reserves Cmd+1), as do the sheet's `formatCells()` and the toolbar's Format Cells… with `sheet`, and OK formats the selection through the sheet's `format`. The dialog is also exported as `FormatDialog`, for a host that places its own. It is Excel's Format Cells for the selection reported by the sheet's `state` event, with the tabs whose settings xlcalc's cell style supports:

- Number is Excel's Number Format box: the category of the code in the input, the list's items (an item puts its format's code in the input) and the code itself, which may be any code.
- Alignment has the horizontal and vertical alignment, the indent (with Left, Right and Distributed alignment, as in Excel) and text wrapping.
- Font has the font name, Regular, Italic, Bold or Bold Italic, the size, the underline, strikethrough and the color.
- Border has the line style and color, the presets None, Outline and Inside, and a button for each edge around a preview, the inside ones for a selection of more than one row or column.
- Fill has the background color.

Colors come from Excel's palette, as on the toolbar, or can be any `#RRGGBB`. The dialog opens with the active cell's formatting; for a single cell the edges show its own borders, and for a range they show as unknown, grayed out as in Excel. OK emits `apply` with patches for `SheetView.format` covering only what changed: the font, the alignment, the fill and the number format over the selection, and each border edge on its row or column, or on the cells on both sides of an inside line (a single patch over every cell when all edges use the same line), all in one call and therefore a single undo step; then it emits `close`, as Cancel, Esc and a click outside the dialog do. Tab and Shift+Tab stay inside it. An input beyond Excel's limits (a size of 1 to 409 points, an indent of 0 to 250) or an invalid color code blocks OK. On a phone the dialog takes the full screen with 44 px targets.

A host with its own dialog listens to the toolbar's `formatCells` event, which the button then emits instead (a host whose listener the toolbar cannot see, such as a custom element, also sets `hostDialog`):

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
