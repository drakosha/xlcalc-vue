# Find and Replace

If the source has `find(req, signal)`, Ctrl+F or the magnifier button on the formula bar opens Excel's Find bar below the formula bar. The sheet sends the source a `FindRequest`: the query with Excel's wildcards (`*` for any characters, `?` for one character, `~` before either of them or before itself to match that character literally) and the options of Excel's dialog under the bar's Options button: Match case, Match entire cell contents, Look in Formulas (the text a cell opens with for editing, Excel's default) or Values (the displayed text), and Search By Rows or By Columns. A selection of more than one cell when the bar opens, or one made while it is open, limits the search (`range`) and is shown on the bar until the user removes it; under a display filter only its visible rows are searched. Enter and Shift+Enter, or the arrow buttons, go to the next and the previous match from the active cell in the chosen order, wrapping around at the end, as Find Next does, and a new query or option goes to its first match after the active cell. A source returns a few hundred matches at a time; beyond them, Find Next requests the next ones from the active cell (`after`, Excel's Find After), so it reaches every match, and the bar then shows the count without the current position. Find Previous goes back over the matches already received. Matches in the viewport are highlighted, and the current one pulses.

`memorySource` has both: press Ctrl+F for Find and Ctrl+H for Replace, then try `?ar*`, or Replace All of `Paris` with `Lyon`.

```vue
<script setup lang="ts">
import { SheetView, memorySource } from '@xlcalc/vue'
import '@xlcalc/vue/style.css'

const cities = ['Paris', 'Berlin', 'Madrid', 'Rome', 'Paris', 'Vienna', 'Warsaw', 'Paris']
const cells = Object.fromEntries(cities.flatMap((city, i) => [[`A${i + 1}`, city], [`B${i + 1}`, (i + 1) * 100]]))
const source = memorySource({ sheets: [{ name: 'Offices', cells }] })
</script>

<template>
  <SheetView :source="source" style="height: 320px" />
</template>
```

## Replace

If the sheet is editable and the source has `replace(req, by)`, Ctrl+H (on a Mac as well, as in Excel for Mac) opens the bar with its Replace row; the row's toggle is at the left of the bar. Replace replaces the active cell if it is a match and goes to the next one; otherwise it only finds, as Excel's does. Replace All replaces every match as a single undo step. Both search in formulas whatever the bar's Look in setting, as Excel's Replace does, and their outcome is handled like an edit's: the changes are drawn, and a rejection is displayed in the sheet's message line. The bar then reports how many cells changed and counts the matches again; cells the source left unchanged (`refused`) are reported in the sheet's message line with the first one's reason.

`memorySource` matches and replaces a cell's text as xlcalc's Find and Replace do (`findMatcher` and `findReplacer` of `xlcalc/text`, measured in Excel): `?` matches one UTF-16 code unit, so a character outside the BMP takes two; without Match case, σ, ς and Σ are the same letter; a trailing `~` is ignored; and in Replace each `*` matches as little as possible, except the last one in the query, which matches the rest. Unlike `localSource`, it searches hidden rows and columns, returns matches in plain row or column order, and Replace All skips cells whose text would not change.
