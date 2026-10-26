# Conditional formatting

The sheet draws the effect of a workbook's conditional formatting rules as its windows give it (`appearance`: a rule's fill, font and borders, data bars, icons). With a source that has `rules(sheet)` and `setRules(sheet, rules)`, on an editable sheet, it also creates, edits and clears the rules with Excel's commands from Home > Conditional Formatting. The sheet has no ribbon, and Excel's cell context menu has no conditional formatting commands, so they are `SheetViewApi` methods for a host's ribbon or menu:

| Excel's command | Method |
| --- | --- |
| Highlight Cells Rules: Greater Than, Less Than, Between, Equal To, Text that Contains, A Date Occurring, Duplicate Values | `quickRule('greaterThan')` … `quickRule('duplicateValues')` |
| Top/Bottom Rules: Top 10 Items, Top 10%, Bottom 10 Items, Bottom 10%, Above Average, Below Average | `quickRule('top10')` … `quickRule('belowAverage')` |
| Data Bars, Gradient Fill and Solid Fill | `dataBars(color, 'gradient' \| 'solid')` |
| Color Scales | `colorScale(preset)` |
| Icon Sets | `iconSet(name)` |
| New Rule… | `newRule()` |
| Clear Rules from Selected Cells, Clear Rules from Entire Sheet | `clearRules('selection')`, `clearRules('sheet')` |
| Manage Rules… | `manageRules()` |

Without `rules` or `setRules`, on a read-only sheet or without a selection, the methods that open a dialog return `false` and the others `null`. `@xlcalc/vue/local` has both source methods; `memorySource` has neither.

```vue
<script setup lang="ts">
import { onUnmounted, ref } from 'vue'
import { SheetView, type SheetViewApi } from '@xlcalc/vue'
import { localSource } from '@xlcalc/vue/local'
import '@xlcalc/vue/style.css'

const source = localSource({
  sheets: [{ name: 'Sales', cells: { A1: { value: 120 }, A2: { value: 80 }, A3: { value: 45 }, A4: { value: 160 } } }],
})
onUnmounted(() => void source.close())
const sheet = ref<SheetViewApi | null>(null)
</script>

<template>
  <button @click="sheet?.quickRule('greaterThan')">Greater Than…</button>
  <button @click="sheet?.dataBars('blue')">Blue data bars</button>
  <button @click="sheet?.colorScale('greenYellowRed')">Green - Yellow - Red</button>
  <button @click="sheet?.iconSet('3Arrows')">3 Arrows</button>
  <button @click="sheet?.clearRules('selection')">Clear Rules from Selected Cells</button>
  <button @click="sheet?.manageRules()">Manage Rules…</button>
  <SheetView ref="sheet" :source="source" :locale="source.locale" style="height: 300px" />
</template>
```

## How a change is written

`setRules` replaces the sheet's list of conditional formats, so each command reads the list with `rules` first and sends the whole new list in one call, a single undo step. A new rule goes first in the list, with the highest priority, as Excel adds one. After a success the sheet reads again the windows over the ranges of the rules before and after the change, as on a `repaint` event, since the outcome carries values only; a source need not send an event for it.

A rule's formulas are in the source's locale, as `rules` returns them (`СУММ($A$1;1)` and `1,5` in ru-RU). The dialogs show them with `=`, as Excel's boxes do, and relative to the active cell: Excel writes a rule's formulas for the top-left cell of its ranges, and so does the sheet, moving the relative references of what was typed from the active cell to that cell. With A2:A3 selected from A3, `=$A3>2` is written as `$A2>2`.

## Quick rules

`quickRule(kind)` opens the small dialog of the command for the selection: a label as Excel's (`Format cells that are GREATER THAN:`), its inputs, and the "with" box of formats. OK writes the rule over the selection and closes the dialog; Cancel and Esc close it.

| Command | Inputs | Rule |
| --- | --- | --- |
| Greater Than, Less Than, Equal To | a value box | `cellIs` with `greaterThan`, `lessThan`, `equal` |
| Between | two value boxes | `cellIs` with `between` |
| Text that Contains | a text box | `containsText` |
| A Date Occurring | Yesterday, Today, Tomorrow, In the last 7 days, Last week, This week, Next week, Last month, This month, Next month | `timePeriod` |
| Duplicate Values | Duplicate or Unique | `duplicateValues` or `uniqueValues` |
| Top 10 Items, Bottom 10 Items | a rank from 1 to 1,000, 10 at first | `top10`, `bottom` for the lowest |
| Top 10%, Bottom 10% | a percentage from 1 to 100, 10 at first | `top10` with `percent` |
| Above Average, Below Average | none | `aboveAverage`, `below` for Below Average |

A value box takes what Excel's takes: a formula after `=` (a reference such as `=$B$1`, checked as a cell's formula and refused with its message), a number, a date, a time or a percentage, stored as its number (`1/1/2020` is 43831, read as a cell reads typed text, `xlcalc/entry`), TRUE or FALSE, or text, stored in quotes (`"done"`). An empty box is refused. A text that begins with `=` is a reference or a formula: the rule gets the formula Excel writes for it (`NOT(ISERROR(SEARCH($A$1,B2)))`); any other text is searched for as typed, by the formula xlcalc gives the rule.

The "with" box offers Excel's formats, the first one chosen:

| Format | Font color | Fill | Border |
| --- | --- | --- | --- |
| Light Red Fill with Dark Red Text | `#9C0006` | `#FFC7CE` | |
| Yellow Fill with Dark Yellow Text | `#9C5700` | `#FFEB9C` | |
| Green Fill with Dark Green Text | `#006100` | `#C6EFCE` | |
| Light Red Fill | | `#FFC7CE` | |
| Red Text | `#9C0006` | | |
| Red Border | | | thin `#9C0006` around the cell |

Custom Format… opens the rule's Format Cells: the Font tab (font style, underline, strikethrough and color: the font's name and size are not part of a rule), the Border tab over a single cell (its four sides) and the Fill tab, all of them the Format Cells dialog's parts. What is left as it opened is not part of the rule; Clear empties every tab. Cancel there returns the "with" box to the preset chosen before.

## Data bars, color scales, icon sets

The galleries' rules are added without a dialog, over the selection.

A data bar is Excel 2010's, with automatic bounds (`autoMin`, `autoMax`), red negative bars (`#FF0000`) and a black axis. Its colors are those of Excel's gallery, the same for both fills: `blue` `#638EC6`, `green` `#63C384`, `red` `#FF555A`, `orange` `#FFB628`, `lightBlue` `#008AEF`, `purple` `#D6007B`. A gradient bar has a border of its own color and a red one on negative bars; a solid bar has none.

A color scale goes from the lowest number to the highest, through the 50th percentile on a three-color scale. Its name lists the colors from the highest numbers to the lowest, as Excel's gallery names them; green is `#63BE7B`, yellow `#FFEB84`, red `#F8696B`, white `#FCFCFF`, blue `#5A8AC6`, and the yellow of the two-color scales `#FFEF9C`.

| Preset | Lowest | Midpoint | Highest |
| --- | --- | --- | --- |
| `greenYellowRed`, `redYellowGreen` | red, green | yellow | green, red |
| `greenWhiteRed`, `redWhiteGreen` | red, green | white | green, red |
| `blueWhiteRed`, `redWhiteBlue` | red, blue | white | blue, red |
| `whiteRed`, `redWhite` | red, white | | white, red |
| `greenWhite`, `whiteGreen` | white, green | | green, white |
| `greenYellow`, `yellowGreen` | yellow, green | | green, yellow |

An icon set takes any of xlcalc's twenty `IconSetName`s, with Excel's thresholds: equal shares in percent, 0, 33 and 67 for three icons, 0, 25, 50 and 75 for four, 0, 20, 40, 60 and 80 for five.

## Clear Rules

`clearRules('sheet')` removes every rule of the sheet. `clearRules('selection')` takes the selection out of each rule that applies to one of its cells: what is left of a range becomes the rows above and below the selection across the range, then the parts to its left and right (`A1:A10` without `A3:A5` is `A1:A2 A6:A10`), and a rule with nothing left is removed. A rule whose top-left cell moved has its formulas written for the new one, so each cell left keeps its condition. Both return `null` and send nothing when there is no rule to clear.

## The Rules Manager

`manageRules()` opens Excel's Conditional Formatting Rules Manager over the sheet's rules, read with `rules`:

- Show formatting rules for: Current Selection (the rules that apply to a cell of the selection, as it opens) or This Worksheet;
- a row for each rule in priority order: its description (`Cell Value > 5`, `Formula: =$A1>5`, `Top 10%`, `Graded Color Scale`), a preview of its format, Applies to as Excel writes it (`=$A$1:$A$10,$C$2`, with the locale's list separator), which can be typed over, and Stop If True, unavailable for data bars, color scales and icon sets;
- New Rule… and Edit Rule… (also a double click) open the rule dialog below; Delete Rule; Duplicate Rule puts a copy right after the rule; Move Up and Move Down change places with the next rule in the list shown.

Applies to takes ranges on this sheet, with or without `=` and `$`; until it does, it is marked and OK and Apply wait. A rule's formulas stay as written when Applies to changes. OK writes the list and closes, Apply writes it and reads it again, and Close discards what was not applied. Edits in the manager are sent only by OK or Apply, in one call each.

## New Formatting Rule and Edit Formatting Rule

The dialog has Excel's six rule types, each with its description and, except the first, a preview and Format… (the rule's Format Cells above):

- Format all cells based on their values: a 2-Color Scale or a 3-Color Scale (each point's type, value and color; a new rule's two-color scale goes from `#FF7128` to `#FFEF9C`, its three-color scale from red through yellow to green), a Data Bar (its minimum and maximum, gradient or solid fill and its color, border or none and its color, Show Bar Only, bar direction), Icon Sets (the icon style, Reverse Icon Order, Show Icon Only, and each icon's threshold: `>=` or `>`, a value and its type);
- Format only cells that contain: Cell Value with its operator and one or two value boxes, Specific Text (containing, not containing, beginning with, ending with), Dates Occurring, Blanks, No Blanks, Errors, No Errors;
- Format only top or bottom ranked values: Top or Bottom, the rank, and % of the selected range;
- Format only values that are above or below average: above, below, equal or above, equal or below, and 1, 2 or 3 standard deviations above or below;
- Format only unique or duplicate values;
- Use a formula to determine which cells to format.

A threshold's value is a number or a formula after `=`; a relative reference there is refused with Excel's message, and a percent or a percentile must be from 0 to 100. An edited rule keeps its Stop If True and what the dialog does not show: a data bar's negative bars, axis and lengths, an icon set's custom icons while its style stays the same.

## Texts and styles

The dialogs' texts are the `cf…` keys of `SheetMessages` (declared in its `ConditionalMessages`), in English and in Russian with Excel's names, and `ok`, `cancel`, `close` and the Format Cells keys (`formatCells`, `fmtTab`, `fmtFontStyle` and the others) for the rule's Format Cells. The dialogs use the sheet's `--xs-*` variables; their classes begin with `xs-cf-`.

## What is left out

- Excel's dialogs show the rule on the cells while it is edited; the sheet shows it after OK.
- Excel fills a quick dialog's value box from the selection's values; the sheet's boxes open empty.
- The Rules Manager lists the rules of the selection or of the displayed sheet, not of another sheet; it has no table or PivotTable rules (Clear Rules from This Table and This PivotTable are not offered).
- The rule's Format Cells has no Number tab: a rule's number format, set elsewhere, is kept.
- The rule dialog does not edit a data bar's negative bar and axis settings or an icon's own icon; an edited rule keeps them.
