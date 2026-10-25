# Formula AutoComplete

A source with `functionNames(signal)` returns the functions a formula can call, by their names as typed in the source's locale (`СУММ` in ru-RU), in the order the list displays them. While a formula is being typed in the cell editor or the formula bar and the caret is right after the beginning of a name (as `nameAt` from `xlcalc/text` finds it: letters, digits, `.` and `_` after the formula's `=`, an operator, `(`, the workbook's list separator (`,` in en-US and `;` in ru-RU), a space or a line break; not inside a string, a quoted sheet name or a table's brackets, nor after a sheet's `!` or a `$`), Excel's list opens below the editor: the functions whose names begin with what was typed, case-insensitively, eight at a time with a scroll bar, the first one highlighted, each with Excel's fx icon; a source with `names` adds the defined names visible on the sheet among them, with a tag icon ([names.md](names.md)). Typing further narrows it, and a character that cannot be part of a name, or a name nothing begins with, closes it.

- The arrow keys, Page Up and Page Down, Home and End move the highlight.
- Tab or a double click replaces the typed name with the function and its `(` (keeping a `(` already after the name), or with the defined name alone, and closes the list.
- Esc closes the list and keeps the editor open, and a second Esc cancels the edit as usual.
- Enter commits the cell as typed, as Excel for Windows does; Excel for Mac also enters the highlighted item on Return.
- Left and Right move the caret, and the list follows the name before it.
- On a touch screen the items are 44 px, and a tap enters one.

The sheet requests the function names once per source, when an editor first opens, and again for a new source; a failed request is retried when the next editor opens, and without `functionNames` (and without defined names) there is no list.

`memorySource` takes the list from its workbook's `functions`: type `=SU` into a cell.

```vue
<script setup lang="ts">
import { SheetView, memorySource } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({
  functions: ['AVERAGE', 'COUNT', 'COUNTIF', 'IF', 'INDEX', 'MATCH', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT'],
  sheets: [{ name: 'Sheet1', cells: { A1: 3, A2: 4 } }],
})
</script>

<template>
  <SheetView :source="source" style="height: 300px" />
</template>
```

The list is `FunctionList`, placed in the grid's content below the cell editor (above it where the viewport has no room below) or below the formula bar's input, and kept horizontally within the viewport; its state and keys are handled by `useAutocomplete` (`src/core/autocomplete.ts`), which the sheet's editing creates. Its texts are `SheetMessages`' `functionList` and `functionHint`, and its colors `--xs-fl-line` and `--xs-on-accent`. `./local` returns every function xlcalc calculates, by its name in the workbook's locale, or in English where the locale has no name of its own (`ABS`, `SIN` in ru-RU, as Excel displays them), sorted in the locale's alphabetical order. `memorySource` returns a workbook's `functions` as given, and has no `functionNames` without them.

## Function argument tooltip

A source with `functionSignature(name, signal)` returns a function's arguments for Excel's argument tooltip (Function ScreenTip): given the name as typed, in English or in the source's locale, in any letter case, it returns `FunctionSignature` (`xlcalc/contract`: the name in the locale, `args` with each argument's name and whether it is optional, `repeat`, the number of trailing arguments that repeat beyond the list, and the tooltip's `text`), or null for a function the source does not know. While the caret is inside a function's parentheses in the cell editor or the formula bar, the tooltip appears below the editor (above it where the viewport has no room below), as Excel's does: `SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`, or `СУММ(число1;[число2];...)` in ru-RU, with the argument at the caret in bold. In a nested call it shows the inner call; beyond the list, the repeating group continues with its numbers counted up (`SUM(number1, [number2], [number3], ...)` at the third argument), and beyond the last argument of a function that takes no more, nothing is bold. There is no tooltip inside a string, after the call's closing parenthesis, outside a formula, or for a name the source does not know. It follows the caret as text is typed or the caret is moved by a key or a click; it closes when the editor closes, on Esc (which cancels the edit, as in Excel) or on commit. The function list takes precedence: while the list is open, the tooltip stays hidden under it, as in Excel, where the two are never displayed together. Unlike Excel's, the function name in the tooltip is not a link to Help. In a narrow viewport the tooltip wraps to fit.

The call and the argument at the caret are found by xlcalc's `callAt` (bundled from `xlcalc/text`) in the workbook's locale. The sheet requests a signature by the function's English name in uppercase (`SUM` for a typed `сумм`, as `callAt` returns it), once per name and source, when the caret first enters a call of that function, not on every keystroke; a request for a call the caret has left is aborted (through its `signal`) and its response discarded, and a failed request is retried the next time. Without `functionSignature` there is no tooltip. The tooltip is `FunctionTip`, next to `FunctionList`, with its state in `useFunctionTip` (`src/core/function-tip.ts`); its colors are the list's `--xs-fl-line`, `--xs-bg` and `--xs-fg`. `./local` responds with xlcalc's `functionSignature` in the workbook's locale; `memorySource` has none.

With `localSource` the list contains every function xlcalc calculates, the tooltip shows their arguments, and fx opens Insert Function: type `=SUMIFS(` into C1 and continue past the commas.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({ sheets: [{ name: 'Sheet1', cells: { A1: { value: 'North' }, B1: { value: 10 }, A2: { value: 'South' }, B2: { value: 20 } } }] })
onUnmounted(() => void source.close())
</script>

<template>
  <SheetView :source="source" :locale="source.locale" style="height: 300px" />
</template>
```

## Insert Function

The fx button to the left of the formula bar's input opens Excel's Insert Function dialog when the sheet is editable and the source has `functionNames`; otherwise the button is disabled. The dialog has a search box, the category box and the list of functions, with the syntax (the signature's `text`) and the description (`description`) of the selected function below it. The search keeps the names that contain the typed text, case-insensitively, with an exact match first, then names that begin with it, then the rest; typing a search while Most Recently Used is selected searches all functions. The category box has Most Recently Used (the ten functions last inserted with OK, newest first; selected by default once there are any, otherwise All), All (the source's `functionNames` in their order) and the categories of the source's signatures (`category`, named by `categoryName`), in the order of Excel's dialog (`FUNCTION_CATEGORIES` of `xlcalc/text`). When it opens, the dialog requests the signature of every name, once per source; the argument tooltip reuses them. The arrow keys move the selection, Enter or a double click means OK, Esc or a click outside the dialog means Cancel, and Tab and Shift+Tab stay inside it. OK starts editing the active cell in the formula bar with `=NAME(` and the caret after the parenthesis, where the argument tooltip continues with its arguments; if an editor is already open, it inserts `NAME(` at the draft's caret (with `=` first in an empty draft), and the edit stays open while the dialog is shown. The recently used functions are kept in the browser's storage under `storageKey`, otherwise for the lifetime of the component. On a phone the dialog takes the full screen, and on a touch screen its inputs and items are 44 px. Its texts are `SheetMessages`' `insertFunction` (the button), `insertFunctionTitle`, `insertFunctionSearch`, `insertFunctionCategory`, `insertFunctionRecent`, `insertFunctionAll` and `insertFunctionPick`, with `ok` and `cancel`.
