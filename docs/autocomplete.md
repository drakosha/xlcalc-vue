# Formula AutoComplete

A source with `functionNames(signal)` gives the functions a formula can call, by their names as the source's locale types them (`СУММ` in ru-RU), in the order the list shows them. While a formula is typed into the cell's field or the formula bar, and the caret stands right after the beginning of a name (letters, digits, `.` and `_` after the formula's `=`, an operator, `(`, an argument separator, a space or a line break; not in a string, a quoted sheet's name or a table's brackets), Excel's list opens under the field: the functions whose names begin with what was typed, in any case, eight at a time with a scroll bar, the first one highlighted. Typing on narrows it, and a character that is no part of a name, or a name no function begins, closes it. The arrows, Page Up and Page Down, Home and End move the highlight; Tab or a double click enters the function in place of the typed name with its `(` (one already after the name is kept) and closes the list; Esc closes the list and keeps the field, and a second Esc cancels the edit as before. Enter commits the cell as typed, as Excel for Windows does; Excel for Mac enters the item on Return too. Left and Right move the caret and the list follows the name before it. Under a finger the items are 44 px and a tap enters one. The sheet asks for the names once per source, when the first field opens, and again for a new source; a request that failed is asked again with the next field, and without `functionNames` there is no list.

`memorySource` gives the list from its book's `functions`: type `=SU` into a cell.

```vue
<script setup lang="ts">
import { SheetView, memorySource } from 'xlcalc-sheet'
import 'xlcalc-sheet/style.css'

const source = memorySource({
  functions: ['AVERAGE', 'COUNT', 'COUNTIF', 'IF', 'INDEX', 'MATCH', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT'],
  sheets: [{ name: 'Sheet1', cells: { A1: 3, A2: 4 } }],
})
</script>

<template>
  <SheetView :source="source" style="height: 300px" />
</template>
```

The list is `FunctionList` in the grid's content under the cell's field (over it where the view has no room below), or under the formula bar's field, kept inside the window; its state and keys are `useAutocomplete` (`src/core/autocomplete.ts`), which the sheet's editing makes. Its texts are `SheetMessages`' `functionList` and `functionHint`, its colours `--xs-fl-*` variables. `./local` gives every function xlcalc computes, by its name in the book's locale, English where the locale names none (`ABS`, `SIN` in ru-RU, as Excel shows them), sorted by the locale's alphabet. `memorySource` gives a book's `functions` as they are, and has no `functionNames` without them.

## The tip with a function's arguments

A source with `functionSignature(name, signal)` gives a function's arguments for Excel's tip (Function ScreenTip): the name as typed, in English or the source's locale, in any case, answers `FunctionSignature` (`xlcalc/contract`: the name in the locale, `args` with each one's name and whether it can be left out, `repeat`, how many of the last arguments go on past the list, and the tip's `text`), or null for a function the source does not know. While the caret stands inside a function's brackets in the cell's field or the formula bar, the tip shows under the field (over it where the view has no room below), as Excel's: `SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`, in ru-RU `СУММ(число1;[число2];...)`, the argument the caret is at in bold. In a call inside a call it is the inner one's; past the list the repeated group goes on with its numbers counted on (`SUM(number1, [number2], [number3], ...)` at the third argument), and past the arguments of a function that takes no more nothing is bold. In a string, after the call's closing bracket, outside a formula or for a name the source does not know there is no tip. It follows the caret as it is typed, moved by a key or a click; it closes when the field closes, by Esc (which cancels the edit, as Excel's) or a commit. The list of functions goes first: while it is open the tip waits under it, as in Excel, where the two never show together. The function's name in the tip is no link to help, as Excel's is. Under a narrow view the tip wraps inside it.

Which call and argument the caret is at is xlcalc's `callAt` (bundled from `xlcalc/text`), in the book's locale. The sheet asks for a signature by the function's English name in upper case (`SUM` for a typed `сумм`, as `callAt` gives it), once per name and source, when the caret first comes into a call of it, not at every key; a request for a call the caret has left is aborted (its `signal`) and its answer dropped, and a failed one is asked again the next time. Without `functionSignature` there is no tip. The tip is `FunctionTip` beside `FunctionList`, its state `useFunctionTip` (`src/core/function-tip.ts`); its colours are the list's `--xs-fl-line`, `--xs-bg` and `--xs-fg`. `./local` answers with xlcalc's `functionSignature` in the book's locale; `memorySource` has none.

With `localSource` the list holds every function xlcalc computes, the tip shows their arguments, and fx opens Insert Function: type `=SUMIFS(` into C1 and go on past the commas.

```vue
<script setup lang="ts">
import { onUnmounted } from 'vue'
import { SheetView } from 'xlcalc-sheet'
import { localSource } from 'xlcalc-sheet/local'
import 'xlcalc-sheet/style.css'

const source = localSource({ sheets: [{ name: 'Sheet1', cells: { A1: { value: 'North' }, B1: { value: 10 }, A2: { value: 'South' }, B2: { value: 20 } } }] })
onUnmounted(() => void source.close())
</script>

<template>
  <SheetView :source="source" :locale="source.locale" style="height: 300px" />
</template>
```

## Insert Function

The fx button left of the formula bar's field opens Excel's Insert Function window when the sheet is editable and the source has `functionNames`; without them the button is off. The window has a search, the category box and the list of functions, with the syntax (the signature's `text`) and the description (`description`) of the one selected under it. The search keeps the names that hold what is typed, in any case, the name itself first, then those it begins, then the rest; typing over Most Recently Used shows all the functions. The category box has Most Recently Used (the ten functions taken last by OK, newest first; the first choice once there are any, else All), All (the source's `functionNames` in their order) and the categories of the source's signatures (`category`, named by `categoryName`), in the order of Excel's window (`FUNCTION_CATEGORIES` in xlcalc; the sheet keeps its own copy, as xlcalc's root carries the engine). The window asks for every name's signature once per source when it opens; the tip keeps them. The arrows move the selection, Enter or a double click is OK, Esc is Cancel. OK opens the active cell's edit in the formula bar with `=NAME(` and the caret after the bracket, where the tip goes on with its arguments; with a field open it puts `NAME(` at the draft's caret (`=` first in an empty draft) and the edit stays open while the window is up. The recently used live in the sheet's memory under `storageKey`, else as long as the component. On a phone the window takes the screen, and under a finger its fields and items are 44 px. Its texts are `SheetMessages`' `insertFunction` (the button), `insertFunctionTitle`, `insertFunctionSearch`, `insertFunctionCategory`, `insertFunctionRecent`, `insertFunctionAll` and `insertFunctionPick`, with `ok` and `cancel`.
