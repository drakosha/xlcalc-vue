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
