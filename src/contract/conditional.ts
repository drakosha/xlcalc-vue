/**
 * Excel's Conditional Formatting menu on the Home tab, as `SheetViewApi` takes its commands: the quick rules with their
 * small dialogs, the galleries of data bars, color scales and icon sets, Clear Rules, New Rule and Manage Rules.
 */

/**
 * A quick rule from Highlight Cells Rules (Greater Than, Less Than, Between, Equal To, Text that Contains, A Date Occurring,
 * Duplicate Values) or Top/Bottom Rules (Top 10 Items, Top 10%, Bottom 10 Items, Bottom 10%, Above Average, Below
 * Average), each opening Excel's small dialog for the selection.
 */
export type QuickRule =
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'equalTo'
  | 'textContains'
  | 'dateOccurring'
  | 'duplicateValues'
  | 'top10'
  | 'top10Percent'
  | 'bottom10'
  | 'bottom10Percent'
  | 'aboveAverage'
  | 'belowAverage'

/** A data bar of Excel's gallery: its color, with a gradient or a solid fill. */
export type DataBarColor = 'blue' | 'green' | 'red' | 'orange' | 'lightBlue' | 'purple'

/** A color scale of Excel's gallery, named from the highest numbers' color to the lowest's (`greenYellowRed`: the highest green). */
export type ColorScalePreset =
  | 'greenYellowRed'
  | 'redYellowGreen'
  | 'greenWhiteRed'
  | 'redWhiteGreen'
  | 'blueWhiteRed'
  | 'redWhiteBlue'
  | 'whiteRed'
  | 'redWhite'
  | 'greenWhite'
  | 'whiteGreen'
  | 'greenYellow'
  | 'yellowGreen'

/** Excel's Clear Rules: from the selected cells, or from the entire sheet. */
export type ClearRulesFrom = 'selection' | 'sheet'
