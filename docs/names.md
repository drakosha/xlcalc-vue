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

## Formula AutoComplete

Formula AutoComplete lists the names visible on the active sheet together with the functions ([autocomplete.md](autocomplete.md)), as Excel lists defined names: the names that begin with the typed text, merged in among the functions in alphabetical order (a name spelt as a function follows it), each with a tag icon where a function has fx. Tab or a double click inserts a name without a parenthesis. Names are offered without the source's `functionNames` too.

## When the names are read

The sheet reads the names once per source, when the workbook loads, and again after each change: after `setNames` (with the sheet, whose formulas may now calculate differently, unless the source sent `reload` meanwhile), after an undo or a redo, and on a `reload` event. The list stays until the new one arrives, and an earlier request's answer is dropped. The Name Box, the Name Manager and AutoComplete all use this one list (`useNames` in `src/core/names.ts`, `names` of the sheet's context).

## Sources

`memorySource` keeps a workbook's `names` as text and checks a name as Excel checks one, without calculating it. `./local` returns its session's names with their formulas translated into the workbook's locale, and parses the formulas of `setNames` in it, refusing one that does not parse; the sheet is read again after a change, since a name changes the values of formulas anywhere.

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
