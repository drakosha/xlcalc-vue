# Defined names

A source with `names(signal)` returns the workbook's defined names as Excel's Name Manager lists them: `SheetName` with the name, what it refers to (`ref`, a formula in the source's locale, without `=`) and the sheet it belongs to (`scope`; a name without one belongs to the workbook). With `setNames(names)` the sheet also defines and changes them: each change sends the entire list, which the source replaces as a single undo step, and a name the source rejects (a name Excel does not accept, one that already exists in its scope, a formula that does not parse) is refused with its message. Without `names` the sheet knows no names; without `setNames`, or on a read-only sheet, it shows them but does not change them.

On a sheet, as in Excel, the names visible are the sheet's own names and the workbook's names, a sheet's own name hiding a workbook name of the same name. A name refers to a range when its formula is a single reference to a range with its sheet (`Sheet1!$B$2:$D$9`, `'Plan 2'!$C:$C`); a constant (`0.2`), a formula, or several areas do not.

## The Name Box

- A defined name typed into the Name Box and Enter go to its range, on its sheet. A sheet's own name of another sheet is typed with its sheet: `'Plan 2'!Total`. A name that does not refer to a range gives Excel's "Reference isn't valid." on the message line; Excel would go to the range a formula such as `OFFSET(...)` returns, which the sheet cannot calculate.
- When the selection is exactly a name's range on the active sheet, the Name Box shows the name instead of the address, as Excel's does (the first such name in alphabetical order).
- The arrow next to it lists the names visible on the active sheet that refer to ranges, in alphabetical order, as Excel's dropdown does; a click on one goes to its range. Alt+Down opens and closes the list from the box, Up, Down, Home and End move in it, Enter goes, Esc closes it.
- A new name typed with a selection (text that is neither an address nor a name) defines it over the selection, in the workbook's scope, as Excel's Name Box does: the name refers to the selection as an absolute reference with its sheet (`Sheet1!$A$1:$A$3`). A name the source refuses gives Excel's "You must enter a valid reference or type a valid name for the selection."

## Define Name

The sheet's `defineName()` opens Excel's New Name dialog (Formulas > Define Name), as does Define Name… in the cell's context menu, and Ctrl+L in Excel for Mac. Its Name is empty, its Scope is Workbook (or one of the sheets, except very hidden ones), and Refers to is the selection as an absolute reference with its sheet and `=` (`=Sheet1!$B$2:$C$3`). OK adds the name through `setNames`; a refusal is shown in Excel's alert with the source's message, and the dialog stays open. The dialog's Comment is not offered: `SheetName` has none. Refers to is typed; selecting a range on the grid while the dialog is open is not supported. `defineName()` returns false when names cannot be defined (no `names` or `setNames`, a read-only sheet) or while a cell is being edited.

## The Name Manager

Ctrl+F3 (on a Mac too) and the sheet's `nameManager()` open Excel's Name Manager: every name of the workbook, sorted by name, with its Refers To (with `=`) and Scope (the sheet's name, or Workbook). Excel's Value column is left out: it would need the source to calculate a name's value. Up and Down move the selection, Enter or a double click edits the selected name.

- New… opens the New Name dialog, Edit… the Edit Name dialog for the selected name (the same fields, with Scope fixed, as Excel's Edit Name has it), and Delete removes it after Excel's question.
- Refers to at the bottom shows the selected name's formula. Editing it enables Commit (the check mark), which replaces the formula, and Cancel (the cross), which restores it; Enter and Esc in the box do the same. Closing with an edit not committed asks whether to save it, as Excel does.
- A refusal is shown in Excel's alert with the source's message; the edit stays for fixing.
- Without `setNames`, or on a read-only sheet, it only lists the names.
- It is not opened while a cell is being edited (`nameManager()` returns false), nor without the source's `names`.

## Create from Selection

Excel's Create from Selection (Formulas > Defined Names) opens the Create Names from Selection dialog for the selection with Ctrl+Shift+F3, Cmd+Shift+F3 on a Mac (as in Excel for Mac), and through the sheet's `createNames()` for a ribbon. Create names from values in the: Top row, Left column, Bottom row, Right column. The dialog opens with the boxes Excel for Mac checked on 17 selections: Top row when the selection has two rows or more and the cell right of its top-left cell (the top-left cell itself in a single column) holds text, a formula's text included; Left column likewise by the cell below the top-left cell (the top-left cell itself in a single row); Bottom row and Right column never. So a table with a header row and a column of row labels opens with both, a header of years (numbers) with none, and a single cell with none. OK is off while no box is checked.

OK closes the dialog and makes the names through the source's `namesFromSelection` (`./local`: xlcalc 0.66.0's `getNamesFromSelection`, which makes them as Excel does: `Q1` becomes `Q1_`, `2025 total` becomes `_2025_total`, a label without text makes none), workbook names referring absolutely to their cells (`P!$M$2:$M$3`). They go to `setNames` in one call with the workbook's other names, so a single undo removes them all. A name the workbook already has with another definition brings Excel's question, `Replace existing definition of 'Q1_'?`: Yes replaces it, No keeps the old one, Cancel stops there, and the names decided before it are made, as Excel makes them one by one. A name with the same definition is left as it is without a question, as Excel for Mac does. Labels with no cell past them (a single row with Top row checked) bring Excel's message, `This selection isn’t valid. Make sure the copy and paste areas don’t overlap unless they are the same size and shape.`, which is the one Excel for Mac shows there, and nothing is made. A refusal of `setNames` goes to the message line. `createNames()` returns false without `namesFromSelection`, `names` or `setNames`, on a read-only sheet or while a cell is being edited. Its texts are `createNamesTitle`, `createNamesIn`, `createNamesTop`, `createNamesLeft`, `createNamesBottom`, `createNamesRight`, `createNamesInvalid` and `createNamesReplace(name)`, with `ok`, `cancel`, `validationYes` and `validationNo`; the dialog's labels and the question are Excel for Windows' in English and in Russian (`Создание имен из выделенного диапазона`), and the message for labels alone is Excel for Mac's.

## Formula AutoComplete

Formula AutoComplete lists the names visible on the active sheet together with the functions ([autocomplete.md](autocomplete.md)), as Excel lists defined names: the names that begin with the typed text, merged in among the functions in alphabetical order (a name spelt as a function follows it), each with a tag icon where a function has fx. Tab or a double click inserts a name without a parenthesis. Names are offered without the source's `functionNames` too.

## When the names are read

The sheet reads the names once per source, when the workbook loads, and again after each change: after `setNames` (with the sheet, whose formulas may now calculate differently, unless the source sent `reload` meanwhile), after an undo or a redo, and on a `reload` event. The list stays until the new one arrives, and an earlier request's answer is dropped. The Name Box, the Name Manager and AutoComplete all use this one list (`useNames` in `src/core/names.ts`, `names` of the sheet's context).

## Sources

`memorySource` keeps a workbook's `names` as text and checks a name as Excel checks one, without calculating it; it has no `namesFromSelection`, whose names follow the engine's rules for labels. `./local` returns its session's names with their formulas translated into the workbook's locale, and parses the formulas of `setNames` in it, refusing one that does not parse; the sheet is read again after a change, since a name changes the values of formulas anywhere.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { SheetView, memorySource, type SheetViewApi } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({
  sheets: [{ name: 'Sheet1', cells: { A1: 'Rent', B1: 1200, A2: 'Food', B2: 400 } }],
  names: [{ name: 'Costs', ref: 'Sheet1!$B$1:$B$2' }, { name: 'Rate', ref: '0.2' }],
})
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <button @click="sheet?.defineName()">Define Name…</button>
  <button @click="sheet?.nameManager()">Name Manager</button>
  <SheetView ref="sheet" :source="source" style="height: 300px" />
</template>
```

Type `Costs` into the Name Box and press Enter, or select A1:A2 and type `Items` there to define a name.

The texts are `SheetMessages`' `nameBoxList`, `nameBoxBad`, `nameNotRange`, `menuDefineName`, `nameItem`, `nameNewTitle`, `nameEditTitle`, `nameLabel`, `nameScope`, `nameRefersTo`, `nameWorkbook`, `nameManager`, `nameNew`, `nameEdit`, `nameDelete`, `nameColumn`, `nameCommit`, `nameDeleteAsk` and `nameSaveAsk`, with `ok`, `cancel`, `close`, `validationYes`, `validationNo` and `validationTitle` (the alert's title).
