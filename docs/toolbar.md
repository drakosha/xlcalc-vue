# The formatting toolbar and Format Cells

## `SheetToolbar`

`SheetToolbar` is Excel's formatting toolbar over the same source, in the order of Excel for Windows' ribbon: Undo and Redo, the Font and Font Size boxes with Increase and Decrease Font Size, Bold, Italic, Underline, Borders, fill and font colors from Excel's palette, Top, Middle and Bottom Align, horizontal alignment, Decrease and Increase Indent, Wrap Text, Merge & Center, Excel's Number Format box with Accounting Number Format, Percent Style, Comma Style, Increase Decimal and Decrease Decimal, for a host that displays the dialog ([below](#formatdialog)) Format Cells, and with `sheet` Excel's AutoSum. It formats the selection reported by the sheet's `state` event through `format(sheet, patches)`, as Excel formats a selection: Bold applies on top of each cell's formatting, and Outside Borders is four patches in one call and therefore a single undo step. Its buttons show as pressed according to the active cell's formatting, and clicking a pressed button removes that formatting, as in Excel. The toolbar is shown only when the source has `format`; Undo and Redo appear when the source has `undo` and `redo`, and are disabled while its `history` (the local source's) has no steps.

Borders is Excel's split button: its part applies the border picked last, Bottom Border until another is picked, and is named after it (`border-last`), and its arrow (`borders`, named Borders) opens the menu with Bottom Border, No Border, All Borders and Outside Borders, in Excel's order with Excel's line after Bottom Border. Excel's other items (Top, Left, Right and the thick and double borders, drawing, More Borders...) are not there; the Format Cells dialog's Border tab draws any edge.

With `sheet` and the source's `structure`, the toolbar has Excel's Merge & Center: the button merges the selection and centers its contents, and its dropdown has Merge & Center, Merge Across, Merge Cells and Unmerge Cells, each a single undo step through the sheet's `merge(command)`. The button shows as pressed on a merged cell, and clicking it then unmerges, as Excel's does; on a single cell outside a merged range it does nothing, and it is disabled on a read-only sheet and under a display filter. Before a merge that would keep only one of several values in a range (in a row, for Merge Across), the sheet shows Excel's warning "Merging cells only keeps the upper-left value and discards other values.", with OK on Enter and Cancel on Esc; it counts the cells that contain a value or a formula in its loaded windows, or in the source's windows over the range when the loaded ones do not cover it. If the source rejects the merge, its message is displayed.

With `sheet`, the toolbar has Excel's AutoSum from Home > Editing, a split button like Merge & Center: Σ enters Sum in the active cell through the sheet's `autoSum()`, opening its edit with `=SUM(range)` and the range selected ([editing.md](editing.md#autosum)), and its dropdown lists Sum, then Average, Count Numbers, Max and Min as the source lists them in `autoSumFunctions` (`./local` lists all four), and More Functions…, which opens the sheet's Insert Function dialog (`insertFunction()`) for a source with `functionNames`. Both parts are disabled while `state.canAutoSum` is absent: without the source's `autoSum` and on a read-only sheet. Its texts are `autoSum`, `autoSumMenu`, `autoSumSum`, `autoSumAverage`, `autoSumCount`, `autoSumMax`, `autoSumMin` and `autoSumMore`.

With `sheet` and the source's `structure`, the toolbar has Excel's Insert > PivotTable, which opens the sheet's PivotTable from table or range dialog through `insertPivotTable()` ([pivot-tables.md](pivot-tables.md#insert-pivottable)); it is disabled while `state.canInsertPivot` is absent, on a read-only sheet. Its text is `pivotInsert`.

With `sheet` and the source's `refreshConnections`, the toolbar has Excel's Data > Refresh All, a split button as on Excel's Data tab: the button asks every call outside the workbook again and refreshes the PivotTables through the sheet's `refreshAll()` ([calculation.md](calculation.md#refresh-all)), and its dropdown lists Refresh All and Refresh, which refreshes the PivotTable under the active cell through `refreshPivot()` ([pivot-tables.md](pivot-tables.md#refresh-and-refresh-all)). The button and Refresh All are disabled while `state.canRefreshAll` is absent, on a read-only sheet; Refresh while `state.canRefreshPivot` is absent. Its texts are Excel for Windows' ScreenTip, `refreshAll`, the name with its key (`Refresh All (Ctrl+Alt+F5)`, `Обновить все (Ctrl+Alt+F5)`), over `refreshAllTip`, its description (`Get the latest data by refreshing all sources in the workbook.`, `Получение новейших данных путем обновления всех источников в книге.`), shown under the name as Office shows it; `refreshAllMenu`, the dropdown's button and list, named Refresh All as Excel names its list; `refreshAllItem` and `menuRefresh` (the items, whose tooltips are `refreshAll` and `refreshPivot`, Refresh's with its key, `Refresh (Alt+F5)`).

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

The toolbar formats the selection, every range of a selection of several (`state.areas`) as Excel does. The Font and Font Size boxes show the active cell's font and size, and each takes a typed value or a pick from its list: Enter or a pick applies it to the selection, and Esc, or leaving the box without Enter, restores what it showed (Esc also emits `done`, so the host can return focus to the grid). A cell without its own font or size shows the workbook's default font (`state.font`, from the source's `SheetInfo.font`), and empty boxes when the source does not give one. A browser does not tell which fonts are installed, so the Font box's list is Excel's common fonts (Aptos, Aptos Narrow, Arial, Calibri, Cambria, Consolas, Courier New, Georgia, Segoe UI, Tahoma, Times New Roman, Verdana), each drawn in its own font, and the `fonts` prop replaces it; any name may be typed, and typing completes it with the first font of the list it begins, as Excel's box does. The Font Size box's list is Excel's (8 to 72), and it takes any size from 1 to 409, with a decimal point or comma, rounded to half points (Excel takes sizes in multiples of .5); any other text stays in the box, marked invalid with the message `fmtSizeRange`, until it is corrected or Esc restores the box. The arrows go through the list, putting each item into the box.

Increase and Decrease Font Size step the active cell's size (11 points for a cell without one, the size the grid draws it at) through Excel's list, beyond 72 in steps of 10 up to 409, and below 8 a point at a time down to 1, and set the new size over the selection; at a limit they send nothing. Top, Middle and Bottom Align show as pressed according to the active cell's vertical alignment, Bottom for a cell without one (Excel's default); clicking a pressed one returns the selection to Bottom. Increase Indent adds a level to the active cell's indent and keeps the Left, Right and Distributed alignments, which have an indent, while any other alignment (General, Center and the rest) becomes Left, as in Excel; Decrease Indent takes a level off without changing the alignment, and does nothing at 0 or with an alignment that has no indent. Indents go up to 250 levels. Percent Style sets Excel's `0%` (the built-in format 9) over the selection, as Ctrl+Shift+% on the grid does ([editing.md](editing.md)); the Number Format box then names Percentage. Accounting Number Format and Comma Style set Excel's built-in formats 44 and 43, which display in the workbook's locale (its currency for Accounting). Increase Decimal and Decrease Decimal take the active cell's format, value and column width (`state.numberFormat`, `cellValue`, `columnWidth`) through xlcalc's `stepDecimals`, as Excel for Mac steps them (a General cell takes the decimals it shows, one more or one less; a code takes a decimal in each section), and put the result on the whole selection, as Excel does; when the step changes nothing (General without a number, no decimal to take away) nothing is sent.

The Number Format box works as Excel's does, as measured by xlcalc (`numberFormatCategory` and `numberFormatChoices` of `xlcalc/text`, without the engine): it shows the category of the active cell's format (General if there is none, Custom where no Excel category matches the code), and its list applies Excel's eleven items (General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text), each with the format Excel sets (Percentage is `0.00%`), after which the box shows the new category. `locale` is the workbook's locale, as for `SheetView`: the category names and the Currency item's code follow it, with ru-RU's data loaded as a separate chunk; without it, en-US. With `sheet` (the sheet's template ref), the toolbar formats, undoes and redoes through the sheet's methods ([api.md](api.md#sheetviewapi)), and the sheet sets the selection and displays a rejection itself. Without it, the toolbar calls the source and emits `select` with where the selection goes after an undo or a redo (for the sheet's `focus`) and `error` with the source's rejection. It emits `done` after each call, so the host can return focus to the grid. Its buttons do not take focus from the grid; the font boxes and the Number Format box do, while they are used. It keeps what fits in one row, or in the lines `rows` allows, and moves the rest into a menu under More ([below](#choosing-the-buttons)); below 600 px or on a touch screen every target is 44 × 44 px. Its buttons and boxes have Excel's ScreenTips instead of the browser's `title`: the name and the key (`Bold (Ctrl+B)`, from `SheetMessages`) below the control after the mouse rests on it for 0.4 s, and at once on the next one within 0.6 s, disabled buttons included; a press, a menu opening or the mouse leaving hides it, and touch gets none. The text is each control's `data-tip` (its `aria-label` names it too), and a host's button shows its `label`.

The names of the buttons and boxes, in English and in Russian, are those of Excel for Windows' ribbon (`Font Color`, `Цвет шрифта`; `Выровнять по верхнему краю`; `Переносить текст`; the Font Size box's `Размер`; AutoSum's `Count Numbers`, `Счетчик`), with the key after the name where Excel's ScreenTip gives one. Two keep their own: Redo is `Вернуть (Ctrl+Y)` in Russian, since Excel's name for that button on its Quick Access Toolbar changes with the history and was measured only without steps to redo, and PivotTable Analyze's Options is `PivotTable Options` (`Параметры сводной таблицы`), since the toolbar is one strip without Excel's contextual tab that would tell which Options it is. `wrapText` names Format Cells' Wrap Text box too, which Excel writes in lower case there.

### Choosing the buttons

The `items` prop picks the toolbar's groups and their order, and adds the host's buttons among them. Without it the toolbar draws every group in the default order, `TOOLBAR_GROUPS`, which is Excel for Windows' ribbon: Undo and Redo from its Quick Access Toolbar, Home's groups in Home's order (Clipboard, Font, Alignment, Number, Editing), then the commands of other tabs in the tabs' order. The groups are `undo`, `formatPainter` ([Format Painter](clipboard.md#format-painter), with `sheet`), `font` (the Font and Font Size boxes with Increase and Decrease Font Size), `fontStyle` (Bold, Italic, Underline), `borders`, `colors` (Fill Color, then Font Color), `verticalAlign`, `align` (horizontal alignment, indent and Wrap Text), `merge`, `number` (the Number Format box and its buttons), `formatCells` (after the Number group, where Excel has the group's dialog launcher), `autoSum`, `pivotTable` (Insert > PivotTable, with `sheet` and the source's `structure`), `refreshAll` (with `sheet` and the source's `refreshConnections`), and `pivotAnalyze` and `pivotDesign` (Excel's PivotTable Analyze and Design commands, shown with `sheet` and the source's `pivotCells` while the active cell is in a PivotTable, [pivot-tables.md](pivot-tables.md#pivottable-analyze-and-design)). A group named in `items` still needs what it needs without it (Undo needs the source's `undo`, Merge & Center and AutoSum need `sheet`); an unknown or repeated group is skipped.

A host's button is an object, `{ id, label, icon?, run?, menu?, disabled?, pressed? }`. The toolbar draws it as its own buttons: `label` is the tooltip and the accessible name, `icon` a Vue component (it gets the class `xs-icon`) or SVG markup, drawn at the toolbar's icon size, and without an icon the label is the button's text. A click closes the toolbar's menus and calls `run`; `disabled` and `pressed` are the host's, so a host that passes a reactive array updates them. Buttons next to each other in `items` form one group. Each has the hook `data-tb-host="<id>"`. With `menu`, a list of `{ id, label, run, disabled? }`, the toolbar draws the button's list as its own (in the narrow menu too): without `run` a click opens it, as Excel's Conditional Formatting does, and with `run` the button is split, as AutoSum: its part runs `run`, its arrow (`data-tb-host-menu="<id>"`) opens the list. An item runs and closes the list; its hook is `data-tb-host-item="<id>"`.

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

At any width the toolbar is one row: it keeps in the row the items that fit, in the order `items` gives them, and moves the rest, the host's buttons included, into the menu under a More button at its end, in order, rather than onto a second line: it measures the row with every item and keeps those from the start that fit beside More, so with the default items a phone keeps Undo, Redo and Format Painter (the Font boxes come next and do not fit), and a host's few buttons after `undo` stay in the row. When every item fits there is no More button. It measures again when its width, the items, `rows` or the texts change.

`rows` lets the items take more lines at 600 px and wider. It is 1 by default. With 2 the toolbar is laid out as Excel's ribbon ([below](#two-rows-as-excels-ribbon)). Any other number wraps the row onto that many lines in order, a group whole onto the next line, and only what does not fit on them goes into More, which then ends the last line; a group that starts a line has no separator before it, and `Infinity` wraps every item, with no More button. Below 600 px the toolbar keeps one line whatever `rows` says, since a phone's 44 px targets would take the sheet's height; a host that wants other lines at other widths passes its own `rows` (`:rows="wide ? 2 : 1"`). A group wider than a line goes into More with what follows it, rather than spilling out of the toolbar.

### Two rows, as Excel's ribbon

With `rows: 2` (at 600 px and wider) the toolbar is one sequence of groups, each two rows tall with a separator between groups, laid out as Excel for Windows lays out its Home tab (measured at 1280 px):

- Font: the Font and Font Size boxes with Increase and Decrease Font Size on the first row, Bold, Italic, Underline, Borders, Fill Color and Font Color on the second. The blocks `font`, `fontStyle`, `borders` and `colors` form this one group where they follow each other in `items`.
- Alignment: Top, Middle and Bottom Align on the first row, Align Left, Center, Align Right and the indents on the second, then Wrap Text over Merge & Center, from `verticalAlign`, `align` and `merge`.
- Number: the Number Format box on the first row, its buttons on the second, and Format Cells at the end of the second row, where Excel has the group's dialog launcher, from `number` and `formatCells`.
- The other groups stack their buttons two to a column, the first at the top, as Excel stacks its small buttons: Undo over Redo (as on Excel's Draw tab), Format Painter and AutoSum at the top of their columns, PivotTable Analyze's commands two to a column. PivotTable, Refresh All and Design's Subtotals, Grand Totals, Report Layout and Blank Rows are Excel's large buttons, as tall as both rows; without Excel's text labels they show a larger icon. A host's run of buttons is a group of its own, stacked the same way.

A run of these blocks that would leave a row empty (`fontStyle` alone, say) stacks its buttons instead, and blocks of one Excel group that do not follow each other in `items` are separate groups. When the groups do not fit, whole groups go into More from the end, none split between the row and the menu, and the menu draws their blocks on one line each, as with `rows: 1`. Below 600 px the toolbar keeps one line, and on a touch screen the rows are 44 px tall.

## `FormatDialog`

`SheetView` has its own Excel Format Cells dialog when the source has `format`: Ctrl+1 opens it (on a Mac as well, since the browser reserves Cmd+1), as do the sheet's `formatCells()` and the toolbar's Format Cells… with `sheet`, and OK formats the selection through the sheet's `format`. The dialog is also exported as `FormatDialog`, for a host that places its own. It is Excel's Format Cells for the selection reported by the sheet's `state` event, with the tabs whose settings xlcalc's cell style supports:

- Number is Excel's Number Format box: the category of the code in the input, the list's items (an item puts its format's code in the input) and the code itself, which may be any code.
- Alignment has the horizontal and vertical alignment, the indent (with Left, Right and Distributed alignment, as in Excel), text wrapping, and Orientation: Degrees, a whole number from -90 to 90 (up for positive ones), and Vertical Text, which stacks the letters and sets the degrees aside, as Excel's Text box does. They are xlcalc's `AlignmentStyle.rotation` (0.66.0), a number or `'vertical'`; 0 removes it.
- Font has the font name, Regular, Italic, Bold or Bold Italic, the size, the underline, strikethrough and the color.
- Border has the line style and color, the presets None, Outline and Inside, and a button for each edge around a preview, the inside ones for a selection of more than one row or column.
- Fill has the background color.
- Protection has Locked and Hidden with Excel's note: a cell is locked and its formula shown unless its style says otherwise, and neither matters until the sheet is protected ([protection.md](protection.md)). On a protected sheet both are disabled, since Excel does not change them there.

Colors come from Excel's palette, as on the toolbar, or can be any `#RRGGBB`. The dialog opens with the active cell's formatting; for a single cell the edges show its own borders, and for a range they show as unknown, grayed out as in Excel. OK emits `apply` with patches for `SheetView.format` covering only what changed: the font, the alignment, the fill, the protection and the number format over the selection, and each border edge on its row or column (of every range of a selection of several, `state.areas`), or on the cells on both sides of an inside line (a single patch over every cell when all edges use the same line), all in one call and therefore a single undo step; then it emits `close`, as Cancel, Esc and a click outside the dialog do. Tab and Shift+Tab stay inside it. An input beyond Excel's limits (a size of 1 to 409 points, an indent of 0 to 250) or an invalid color code blocks OK. On a phone the dialog takes the full screen with 44 px targets.

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

Fill patterns, diagonal borders, shrink to fit, merging, text direction, superscript, subscript and the Protection tab are not available, since xlcalc's cell style does not support them. Excel's dial for the degrees is not drawn; the degrees are typed.

### Rotated text

A cell whose style has `alignment.rotation` is drawn as Excel orients its text: at an angle, the line turned counterclockwise by positive degrees and clockwise by negative ones, about the middle of the box it takes when turned, and that box aligned in the cell by the cell's horizontal and vertical alignment as level text is; `'vertical'` stacks the letters one above the other. The box is measured in the cell's font on a canvas (estimated where the browser has none). Unlike Excel, rotated text does not overflow into empty neighbors and does not wrap, borders are not slanted with text at an angle, and a hyperlink's text stays level: the text is clipped to its cell. The row's height and the text are the source's: `./local` (xlcalc 0.67.0) fits a row to its rotated text as Excel for Mac does, when text is typed into a rotated cell, when Format Cells or the toolbar rotates text, and on AutoFit, unless the row's height was set by hand, and the sheet reads the heights again; a rotated cell shows its text as Excel's widest column would, so a number turned in a narrow column shows all its digits rather than `####`.
