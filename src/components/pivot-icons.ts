// The PivotTable icons: Insert > PivotTable on the toolbar and the areas of the PivotTable Fields pane, as strokes on the icons' grid.
import { icon } from './icons'

/** A table with its header row and label column, and lines of values: Excel's PivotTable button. */
export const IconPivotTable = icon(['M3 3h18v18H3z', 'M3 9h18', 'M9 3v18', 'M13 13h5', 'M13 17h5'])
/** The Columns area: a box of columns. */
export const IconColumns = icon(['M3 4h18v16H3z', 'M9 4v16', 'M15 4v16'])
/** The Rows area: a box of rows. */
export const IconRows = icon(['M3 4h18v16H3z', 'M3 9.5h18', 'M3 14.5h18'])
/** The Values area and Σ Values: a sigma. */
export const IconSigma = icon(['M18 5H6l6 7-6 7h12'])
/** PivotTable Options: a table with a gear's tooth. */
export const IconPivotOptions = icon(['M3 3h12v12H3z', 'M3 8h12', 'M8 3v12', 'M18 14v2.5', 'M18 19.5V22', 'M14.5 18h2', 'M19.5 18h2', 'M18 18h.01'])
/** Group Selection: two boxes joined by a bracket. */
export const IconGroup = icon(['M4 4h6v6H4z', 'M4 14h6v6H4z', 'M14 4h3v16h-3', 'M17 12h4'])
/** Ungroup: two boxes and a bracket broken. */
export const IconUngroup = icon(['M4 4h6v6H4z', 'M4 14h6v6H4z', 'M14 4h3v5', 'M17 15v5h-3', 'm19 10 2 4'])
/** Group Field: a box of rows with a bracket. */
export const IconGroupField = icon(['M3 4h10v16H3z', 'M3 9.5h10', 'M3 14.5h10', 'M16 4h3v16h-3'])
/** Fields, Items, & Sets: a function's fx. */
export const IconCalculated = icon(['M9 4H7a2 2 0 0 0-2 2v14', 'M3 11h6', 'm13 10 7 9', 'm20 10-7 9'])
/** Subtotals: rows with a total under a group. */
export const IconSubtotals = icon(['M3 4h18', 'M6 8h15', 'M6 12h15', 'M3 16h18', 'M3 20h18'])
/** Grand Totals: a box with its last row and column drawn. */
export const IconGrandTotals = icon(['M3 3h18v18H3z', 'M3 16h18', 'M16 3v18'])
/** Report Layout: a table with its labels' column. */
export const IconReportLayout = icon(['M3 3h18v18H3z', 'M3 8h18', 'M9 8v13', 'M5 12h2', 'M5 16h2'])
/** Blank Rows: rows with a gap between them. */
export const IconBlankRows = icon(['M3 4h18', 'M3 8h18', 'M3 16h18', 'M3 20h18', 'M6 12h.01', 'M12 12h.01', 'M18 12h.01'])
