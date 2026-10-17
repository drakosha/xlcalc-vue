# Data validation

A source with `validation(sheet, cell, signal)` gives what Excel shows on a cell under a rule of data validation (`SheetValidation`): the items of a list rule's dropdown as their text, in order (empty when the list's formula gives none), and the input message, its title and text; null when no rule is over the cell. An edit or a Control+Return whose value a rule does not take is refused with `alert` (`SheetValidationAlert`): the rule's style, `stop`, `warning` or `information`, its title (none for Excel's "Microsoft Excel"), its message and the cell. After a warning's Yes or an information's OK the component sends the same call again with `{ accept: true }`, and the source writes the value without asking the rule again. A paste is not checked, as Excel does not check one. `memorySource` takes a sheet's `validations` (`MemoryValidation`: a list or whole numbers within bounds, an input message, the alert's style, title and message), enough for an example or a test; `./local` takes the book's rules ([local.md](local.md#data-validation)).

A list rule with an input message and a rule of whole numbers with a warning, over `memorySource`: select a cell of B2:B6 for its arrow (Alt+Down opens the list), and type 150 into C2.

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

The parts that show it are Excel's. `useValidation` (`src/core/validation.ts`) asks the source once the active cell stays for 60 ms (and after the same pause once a stream of the source's events stops), aborts the request of a cell left and drops its answer, and keeps the list's state: Alt+Down opens the list at the item the cell shows, its items asked again as Excel computes them on opening; the arrows, Page Up and Page Down (eight items, as many as the list shows), Home and End move; Enter or a click enters the item into the cell as typed, Esc and Alt+Up close it, and any other key closes it and goes on to the grid. `ValidationDropdown` is the arrow right of the active cell, at its bottom, and the list under the cell as wide as the cell and its arrow, over the cell where the view has no room below; `ValidationPrompt` is the input message, a pale yellow box below the cell from its middle, left of that or over the cell where the view ends; `ValidationAlert` is the alert with its style's icon and buttons: Stop has Retry and Cancel, Warning the rule's message, "Continue?" and Yes, No and Cancel, Information OK and Cancel. It opens on Excel's default button (Retry, No, OK) and Esc is Cancel; Yes and OK keep the value, Retry and No open the cell's field again with the text typed, Cancel leaves the cell as it was. Under a finger the arrow and the items are 44 px, and on a phone the alert's buttons too. Their texts are `SheetMessages`' (`validationList`, `validationTitle`, `retry`, `validationYes`, `validationNo`, `validationContinue`, `validationStyle`, with `ok` and `cancel`), their colours `--xs-prompt-*` and `--xs-alert-*` variables. `SheetView` mounts them: the arrow and the message by the active cell when the source has `validation`, the alert when an edit or a Control+Return is refused with `alert`, and an item enters through the source's `pick(sheet, cell, item, text)`, by its place in the list and its text, which writes it as Excel writes a picked item: its value, not read as typed (a list's cell holding the text `001` gives `001`, one showing `50%` the number 0.5), the cell's format kept, the rule not asked. A source without `pick` (`memorySource`) gets the item's text through `edit` with `accept`. While the alert is up the edited cell is the active one and shows the typed text as pending.
