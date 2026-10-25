# Data validation

A source with `validation(sheet, cell, signal)` returns what Excel displays for a cell under a data validation rule (`SheetValidation`): the items of a list rule's dropdown as text, in order (empty when the list's formula returns none), and the input message with its title and text; null when no rule applies to the cell. An edit or a Ctrl+Enter whose value violates a rule is rejected with `alert` (`SheetValidationAlert`): the rule's style (`stop`, `warning` or `information`), its title (none for Excel's default "Microsoft Excel"), its message and the cell. After Yes on a warning or OK on an information alert, the component sends the same call again with `{ accept: true }`, and the source writes the value without checking the rule again. A paste is not checked, as Excel does not check one. `memorySource` accepts a sheet's `validations` (`MemoryValidation`: a list or whole numbers within bounds, an input message, and the alert's style, title and message), which is enough for an example or a test; `./local` uses the workbook's rules ([local.md](local.md#data-validation)).

A list rule with an input message and a whole-number rule with a warning, over `memorySource`: select a cell in B2:B6 to see its dropdown arrow (Alt+Down opens the list), and type 150 into C2.

```vue
<script setup lang="ts">
import { SheetView, memorySource } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const source = memorySource({
  sheets: [
    {
      name: 'Tasks',
      cells: { A1: 'Task', B1: 'Status', C1: 'Done, %', A2: 'Design', B2: 'Open', C2: 40, A3: 'Build', B3: 'On hold', C3: 0 },
      validations: [
        { range: 'B2:B6', list: ['Open', 'On hold', 'Closed'], prompt: { title: 'Status', text: 'Pick one from the list.' } },
        { range: 'C2:C6', whole: { min: 0, max: 100 }, alert: { severity: 'warning', title: 'Done, %', message: 'A whole number from 0 to 100.' } },
      ],
    },
  ],
})
</script>

<template>
  <SheetView :source="source" style="height: 300px" />
</template>
```

## Dropdown, input message and alert

These parts look and behave like Excel's. `useValidation` (`src/core/validation.ts`) requests the rule once the active cell has stayed in place for 60 ms (and after the same pause once a burst of source events ends), aborts the request when the active cell moves away and discards its response, and keeps the list's state:

- Alt+Down opens the list at the item the cell displays and requests the items again, since Excel recalculates them when the list opens;
- the arrow keys, Page Up and Page Down (eight items, as many as the list displays), Home and End move the highlight;
- Enter or a click enters the item into the cell, Esc and Alt+Up close the list, and any other key closes it and goes on to the grid.

`ValidationDropdown` is the arrow at the bottom right of the active cell, and the list below the cell, as wide as the cell and its arrow, or above the cell where the viewport has no room below. `ValidationPrompt` is the input message: a pale yellow box below the cell starting from its middle, shifted left or placed over the cell where the viewport ends. `ValidationAlert` is the alert with its style's icon and buttons: Stop has Retry and Cancel; Warning has the rule's message, "Continue?" and Yes, No and Cancel; Information has OK and Cancel. It opens with Excel's default button focused (Retry, No, OK), and Esc means Cancel. Yes and OK keep the value, Retry and No reopen the cell editor with the typed text, and Cancel leaves the cell unchanged. On a touch screen the arrow and the list items are 44 px, and on a phone so are the alert's buttons. Their texts come from `SheetMessages` (`validationList`, `validationTitle`, `retry`, `validationYes`, `validationNo`, `validationContinue`, `validationStyle`, with `ok` and `cancel`), and their colors from the `--xs-prompt-*` and `--xs-alert-*` variables.

`SheetView` mounts these parts: the arrow and the input message next to the active cell when the source has `validation`, and the alert when an edit or a Ctrl+Enter is rejected with `alert`. A chosen item is written through the source's `pick(sheet, cell, item, text)`, by its position in the list and its text, and the source writes it as Excel writes a picked item: the item's value, not parsed as typed text (a list cell containing the text `001` gives `001`, one displaying `50%` gives the number 0.5), keeping the cell's number format and without checking the rule. A source without `pick` (`memorySource`) receives the item's text through `edit` with `accept`. While the alert is open, the edited cell stays active and displays the typed text as pending.

## The Data Validation dialog

`SheetViewApi.dataValidation()` opens Excel's Data Validation dialog (Data > Data Validation) for the selection, for a host's ribbon or menu: Excel's cell context menu has no entry for it, so the sheet adds none. It needs the source's `rules` and `setRules` and an editable sheet, and returns `false` otherwise. `memorySource` has neither method; `./local` has both.

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { SheetView, type SheetViewApi } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({
  sheets: [
    {
      name: 'Tasks',
      cells: { A1: { value: 'Status' }, D1: { value: 'Statuses' }, D2: { value: 'Open' }, D3: { value: 'On hold' }, D4: { value: 'Closed' } },
      validations: [{ range: 'A2:A20', type: 'list', formula1: '$D$2:$D$4' }],
    },
  ],
})
onUnmounted(() => void source.close())
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <button @click="sheet?.dataValidation()">Data Validation…</button>
  <SheetView ref="sheet" :source="source" :locale="source.locale" style="height: 300px" />
</template>
```

When it opens, the sheet reads the sheet's rules from `rules` and looks at the selection as Excel does:

- no rule in the selection: the dialog opens with Excel's defaults (Any value);
- one rule's settings over every cell: the dialog opens with them;
- the same settings over some cells only: Excel's question "The selection contains some cells without Data Validation settings. Do you want to extend Data Validation to these cells?", where Yes opens the dialog with those settings for the whole selection, No opens it with the defaults, and Cancel opens nothing;
- rules of more than one kind: "The selection contains more than one type of validation. Erase current settings and continue?", where OK opens the dialog with the defaults and Cancel opens nothing.

Rules count as one kind when every setting is the same, their formulas compared as written for their own cells.

The dialog has Excel's three tabs. Settings has Allow (Any value, Whole number, Decimal, List, Date, Time, Text length, Custom), Data with Excel's eight comparisons, the boxes the type and the comparison call for (Minimum and Maximum, Value, Length, Start date and End date, Start time and End time, Source, Formula), Ignore blank, In-cell dropdown for a list, and "Apply these changes to all other cells with the same settings", offered when the dialog opened with a rule's settings. Input Message has its check box, Title (32 characters, as Excel allows) and Input message (255); Error Alert has its check box, Style (Stop, Warning, Information), Title (32) and Error message (225). Clear All returns every tab to the defaults and leaves the dialog open, so OK then removes the rule from the selection.

The boxes show a rule's formulas in the source's locale, with `=`, written for the active cell (a rule `B2>A2` over B2:B9 shows `=B5>A5` with B5 active); a number shows as it is (`10`, `1,5` in ru-RU), a date or time bound in the locale's short date or time (`1/1/2025`, `01.01.2025`, `9:00:00 AM`), and a literal list with the locale's list separator: `Open,On hold,Closed` in en-US and `Open;On hold;Closed` in ru-RU. OK reads the boxes as Excel reads them: `=` starts a formula, and anything else is read as typed into a cell (`xlcalc/entry`), so a bound must be a number, a date or a time, and a list typed with the locale's separator is stored with commas, as a file stores it in every locale (`"Open,On hold,Closed"`). Before writing, OK refuses an empty box ("You must enter a value in the Source box."), a bound that is not a value ("The Minimum you entered is not valid.") and a maximum below the minimum ("The Maximum must be greater than or equal to the Minimum."), on the Settings tab, and leaves the dialog open.

OK writes the sheet's whole list of data validation rules in one `setRules` call: the new rule over the selection (or none for Any value without an input message or alert text), with "Apply these changes to all other cells with the same settings" over every cell with the settings the dialog opened with too. A rule the selection covers only part of keeps its other cells, with its formulas written again for the top-left cell of what remains, as a rule's formulas are relative to that cell. When the source refuses (a formula that does not parse), the dialog stays open with its message. After the call the sheet re-reads the windows over the changed cells and requests the active cell's rule again, so its dropdown and input message follow at once. The texts are the `SheetMessages` keys `dataValidation` and those starting with `dv`, with the style names of `validationStyle`.

Circle Invalid Data and Clear Validation Circles are not offered: they need each cell's value checked against its rule, and the source contract has no such check (`validation` returns a cell's dropdown and input message, and an edit is checked only as it is written). Excel's warning "The Source currently evaluates to an error" before OK is not given either, since the sheet cannot calculate the list's formula; the source's own check of the formula applies.
