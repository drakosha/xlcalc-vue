# Calculation options

Excel's Formulas > Calculation Options sets how a workbook calculates: Automatic, Automatic Except for Data Tables, or Manual. The workbook keeps the option and Excel saves it with the file. The sheet takes the option from its source (`calculation`), shows Calculate on the status bar while cells wait for a calculation, and calculates with F9 and Shift+F9 through the source's `calculate`. A host changes the option through `SheetViewApi.setCalculation`, for instance from its own Calculation Options button.

## Manual

In Manual, as Excel for Mac 16.114 behaved when measured:

- An edit shows the edited cell and leaves the formulas that depend on it with their values: after typing 5 into `A1`, `=A1*2` in `B1` still shows 2. The status bar shows Calculate.
- A formula entered, pasted or filled in Manual shows its value at once (`=A1*10` shows 50), with the cells its array fills and empties; entering a formula again calculates it the same way.
- A change of how a cell looks, such as a number format, a narrower column or Clear Formats, shows at once with the cell's value as of the last calculation, and nothing starts waiting.
- Undo and redo restore the edited cells and leave the dependents too.
- Calculate Now (F9, or a click on Calculate) shows the values of every waiting cell; Calculate Sheet (Shift+F9) those of the active sheet only, so the status bar keeps Calculate while another sheet waits.
- Switching to Automatic calculates what waits.

Changing the option is not a step of undo here: Excel keeps it as a setting of the application and the workbook, outside the undo list (not measured, since a script's changes do not reach Excel's undo).

Automatic Except for Data Tables calculates as Automatic and keeps the option with the workbook: Excel then leaves only data tables (What-If Analysis > Data Table) for F9, and the workbook model has no data tables.

## Saving

Excel calculates a workbook before saving it (its Recalculate workbook before saving option, on by default), so a file saved in Manual holds the calculated values, and the sheet shows them after the save. `saveXlsx` of `@xlcalc/vue/local` does the same: in Manual it calls the source's `calculate` first, so the values written are the values shown. The file keeps the option (`calcPr calcMode`), and a file opened from it is in the same mode.

## In `@xlcalc/vue/local`

xlcalc records the option (`BookInput.calculation`) and calculates every edit whatever it says. `localSource` reads the option from the workbook and keeps showing the values of the last calculation in Manual: before the first change in Manual it captures the values shown of every formula and of the cells its array fills, and a value that a change recalculated waits in the source, which puts it back over the session's in each window, until `calculate` shows the session's. `book()`, `snapshot()` and so `saveXlsx` carry the option as `setCalculation` left it.

The capture reads every formula of the workbook once per stretch of Manual, and keeps the values shown in memory until the option leaves Manual; a workbook with hundreds of thousands of formulas pays for it on the first change.

```ts
import { localSource, openXlsx, saveXlsx } from '@xlcalc/vue/local'

const { book } = await openXlsx(file) // a file saved in Manual opens in Manual
const source = localSource(book)
await source.edit([{ sheet: 'Sheet1', cell: 'A1', text: '5' }])
console.log(await source.calculation()) // { mode: 'manual', pending: true }
await source.calculate() // F9: the waiting cells show their values
await source.setCalculation('automatic')
```

## Differences from Excel

- A formula entered or pasted in Manual shows the value calculated from its precedents' current values; Excel calculates it from the values they show, so `=B1` over a waiting `B1` shows the old value there.
- A cell that waits keeps the text it showed when its number format or its column's width changes, until a calculation; Excel draws the old value in the new format (not measured).
- Calculate Sheet shows the session's values of the sheet's cells, calculated from every sheet; Excel calculates the sheet over the values the other sheets show.
- A structure operation in Manual (inserting or deleting rows, columns and sheets, moving a range, a table command), and its undo, shows the calculated values of every cell: the values waiting are held by address, which the operation moves. Excel keeps the old values until F9.
- The display filter's lists, Find in values, the selection's totals and Copy read the calculated values, where Excel reads the values shown.
- In Automatic, F9 recalculates nothing: the session's values are current. Excel recalculates volatile functions such as `RAND` and `NOW` on F9.

## A ribbon's Calculation Options

The sheet has no ribbon; a host adds the menu as one of `SheetToolbar`'s buttons, with the option the `state` event reports:

```vue
<script setup lang="ts">
import { computed, ref } from 'vue'
import { SheetToolbar, SheetView, type SheetSource, type SheetViewApi, type SheetViewState, type ToolbarItem } from '@xlcalc/vue'

const props = defineProps<{ source: SheetSource }>()
const sheet = ref<SheetViewApi | null>(null)
const state = ref<SheetViewState | null>(null)
const check = (on: boolean) => (on ? '✓ ' : '')
const items = computed<ToolbarItem[]>(() => [
  'undo',
  {
    id: 'calc',
    label: 'Calculation Options',
    menu: [
      { id: 'auto', label: `${check(state.value?.calculation === 'automatic')}Automatic`, run: () => void sheet.value?.setCalculation('automatic') },
      { id: 'table', label: `${check(state.value?.calculation === 'autoNoTable')}Automatic Except for Data Tables`, run: () => void sheet.value?.setCalculation('autoNoTable') },
      { id: 'manual', label: `${check(state.value?.calculation === 'manual')}Manual`, run: () => void sheet.value?.setCalculation('manual') },
    ],
  },
  { id: 'now', label: 'Calculate Now', run: () => void sheet.value?.calculate() },
])
</script>

<template>
  <SheetToolbar :source="props.source" :state="state" :sheet="sheet" :items="items" />
  <SheetView ref="sheet" :source="props.source" @state="(s) => (state = s)" />
</template>
```
