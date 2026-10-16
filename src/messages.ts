// Interface texts of the component: English by default, a Russian set in the package; a host passes its own or a part.
import { computed, type ComputedRef } from 'vue'

export interface SheetMessages {
  back: string
  backEmpty: string
  address: string
  addressBad: string
  goSheet: (name: string) => string
  showRange: string
  cellLink: string
  cellLinkCopied: (ref: string) => string
  clipboardDenied: string
  /** The lock's title: what the book freezes, and whether it is on now. */
  freeze: (rows: number, cols: number, on: boolean) => string
  note: (author?: string) => string
  linkInBook: (location: string) => string
  linkOpen: string
  linkDead: (href: string) => string
  loading: string
  empty: string
  retry: string
  loadFailed: string
  copyFailed: string
  copyPartial: string
  /** Selection totals over loaded cells: count, then sum, average, min, max and numbers when there are any. */
  statsCount: (n: string) => string
  statsSum: (n: string) => string
  statsAverage: (n: string) => string
  statsMin: (n: string) => string
  statsMax: (n: string) => string
  statsNumbers: (n: string) => string
  /** A selection larger than the loaded windows: only its size. */
  statsCells: (n: string) => string
  sheetTitle: (name: string, rows: number, cols: number) => string
  sheetUnnamed: (n: number) => string
  /** "+" after the tabs, and the word a new sheet's name begins with (Excel's `Sheet4`, `Лист4`). */
  insertSheet: string
  sheetNamePrefix: string
  /** The structure menu on row numbers, column letters and tabs, as Excel names its items. */
  menu: string
  menuInsert: string
  menuDelete: string
  menuColumnWidth: string
  menuRowHeight: string
  menuHide: string
  menuUnhide: string
  menuRename: string
  menuMoveLeft: string
  menuMoveRight: string
  menuUnhideSheet: string
  /** Excel's outline: Group and Ungroup in the menu and its box over cells (rows or columns), the buttons' titles, what is not done. */
  menuGroup: string
  menuUngroup: string
  outlineRows: string
  outlineColumns: string
  outlineExpand: string
  outlineCollapse: string
  outlineLevel: (n: number) => string
  outlineDeepest: string
  outlineNone: string
  /** Excel's menu on cells (cut, copy and paste inside the book), a paste the browser will not read, the fill handle. */
  cellMenu: string
  menuCut: string
  menuCopy: string
  menuPaste: string
  pasteDenied: string
  fillHandle: string
  /** Excel's commands in the cells' menu: Reapply of the display filter, Freeze Panes and its kinds, Trace Precedents as a panel beside the grid. */
  menuReapply: string
  filterReapply: string
  menuFreezePanes: string
  menuUnfreezePanes: string
  menuFreezeTopRow: string
  menuFreezeFirstColumn: string
  menuPrecedents: string
  explainTitle: string
  explainReading: string
  explainFailed: string
  explainEmpty: string
  /** The size boxes: their labels in Excel's units and the limits. */
  columnWidth: string
  rowHeight: string
  widthRange: string
  heightRange: string
  ok: string
  cancel: string
  /** The question before a sheet with data is deleted, and the list of hidden sheets. */
  deleteSheet: string
  deleteSheetText: (name: string) => string
  unhideSheet: string
  rename: string
  selectColumn: string
  selectRow: string
  /** List mode: a sortable column's title, and the first column's cells that open their row. */
  sortColumn: string
  openRow: string
  selectAll: string
  resize: string
  /** A column's edge where a double click fits the column to what it shows. */
  resizeFit: string
  /** The grip on a row number's edge. */
  resizeRow: string
  conflict: string
  locked: (reason: string) => string
  editFailed: string
  /** A paste whose rectangle would go past the sheet's last row or column. */
  pasteOffSheet: string
  tooManyCells: string
  pending: string
  close: string
  /** The formula bar's tip over a formula with references: Ctrl+click (Cmd on a Mac, the key given) goes to one. */
  goRefHint: (key: string) => string
  /** Excel's Enter button left of the formula bar (its Cancel button is `cancel`): the edit entered, the selection kept. */
  barEnter: string
  /** Excel's Insert Function: the fx button, then its window's title, search, category box with its first two items, and the list's label. */
  insertFunction: string
  insertFunctionTitle: string
  insertFunctionSearch: string
  insertFunctionCategory: string
  insertFunctionRecent: string
  insertFunctionAll: string
  insertFunctionPick: string
  /** Excel's button at the formula bar's right end and its Control+Shift+U: the bar one line high or several. */
  expandBar: string
  collapseBar: string
  /** A book being recomputed: the wait, its seconds, and the end of patience. */
  waiting: string
  waitSeconds: (n: number) => string
  waitHint: string
  waitTimeout: string
  find: string
  findPlaceholder: string
  findPrev: string
  findNext: string
  findNone: string
  /** Search status: the hit's place, all hits, and how many came when the source cut the list. */
  findStatus: (i: number, total: number, shown: number) => string
  /** Search status while the active cell is no hit: all hits, and how many came. */
  findCount: (total: number, shown: number) => string
  findFailed: string
  /** Excel's Find and Replace: its options, the range it looks within, the replace row and what Replace All did. */
  findOptions: string
  findMatchCase: string
  findWholeCell: string
  findLookIn: string
  findInFormulas: string
  findInValues: string
  findOrder: string
  findByRows: string
  findByColumns: string
  findWithin: (range: string) => string
  findWholeSheet: string
  replaceToggle: string
  replacePlaceholder: string
  replaceOne: string
  replaceAll: string
  replaceDone: (n: number) => string
  /** Replace All left `n` cells as they were: Excel would not take their new text, for the source's reason `why`. */
  replaceRefused: (n: number, why: string) => string
  filter: string
  filterOn: string
  filterColumn: (letters: string) => string
  filterNamed: (name: string, letters: string) => string
  sortAsc: string
  sortDesc: string
  filterText: string
  filterSearch: string
  filterAll: string
  filterEmpty: string
  filterMore: (n: number) => string
  filterReading: string
  filterApply: string
  filterReset: string
  filterClear: string
  /** The filter line at the bottom: rows shown of all, the columns with conditions, the sort. */
  filterStatus: (shown: string, total: string, cols: string[], sort: { col: string; dir: 'asc' | 'desc' } | null) => string
  filterFailed: string
  valuesFailed: string
  unfiltered: (row: number) => string
  /** The formatting toolbar (`SheetToolbar`): its buttons, menus and number formats as Excel names them. */
  toolbar: string
  toolbarMore: string
  undo: string
  redo: string
  bold: string
  italic: string
  underline: string
  fontColour: string
  fillColour: string
  pickColour: string
  colourAuto: string
  fillNone: string
  borders: string
  borderAll: string
  borderOutside: string
  borderBottom: string
  borderNone: string
  alignLeft: string
  alignCenter: string
  alignRight: string
  wrapText: string
  numberFormat: string
  /** Excel's Format Cells dialog (`FormatDialog`): its tabs, fields and the names of what they offer. */
  formatCells: string
  formatCellsOpen: string
  fmtTab: (tab: 'number' | 'alignment' | 'font' | 'border' | 'fill') => string
  fmtCategory: string
  fmtCode: string
  fmtCodeEmpty: string
  fmtHorizontal: string
  fmtVertical: string
  fmtHorizontalKind: (h: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed') => string
  fmtVerticalKind: (v: 'top' | 'center' | 'bottom' | 'justify' | 'distributed') => string
  fmtIndent: string
  fmtIndentRange: string
  fmtFontName: string
  fmtFontDefault: string
  fmtFontStyle: string
  fmtFontStyleKind: (k: 'regular' | 'italic' | 'bold' | 'boldItalic') => string
  fmtSize: string
  fmtSizeRange: string
  fmtUnderline: string
  fmtUnderlineKind: (k: 'none' | 'single' | 'double') => string
  fmtStrike: string
  fmtColour: string
  fmtColourCode: string
  fmtColourBad: string
  fmtBackground: string
  fmtLine: string
  fmtLineStyle: (style: string) => string
  fmtPresets: string
  fmtPreset: (p: 'none' | 'outline' | 'inside') => string
  fmtEdges: string
  fmtEdge: (e: 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV') => string
  /** Excel's data validation: the dropdown's button, the alert's title without the rule's, its buttons (Retry, OK and Cancel are `retry`, `ok` and `cancel`), a warning's question and the styles' names. */
  validationList: string
  validationTitle: string
  validationYes: string
  validationNo: string
  validationContinue: string
  validationStyle: (severity: 'stop' | 'warning' | 'information') => string
  /** The question on a formula that does not read, under the source's reason, and its buttons. */
  formulaQuestion: string
  formulaFix: string
  formulaRevert: string
  /** Excel's window on a formula left open that it offers to correct (Yes and No are `validationYes` and `validationNo`), and its message after No. */
  typoFound: (correction: string) => string
  missingParenthesis: string
  /** Excel's Formula AutoComplete: the list of functions under the field and its hint. */
  functionList: string
  functionHint: string
  /** Excel's Merge & Center button, its list's button and items, and the warning before values are lost. */
  mergeCenter: string
  mergeMenu: string
  mergeAcross: string
  mergeCells: string
  unmergeCells: string
  mergeWarning: string
}

/** `n` and its English noun, singular for one. */
const countEn = (n: number, one: string, many: string) => `${n} ${n === 1 ? one : many}`
/** The tail of a search status when the source cut the list. */
const shownEn = (total: number, shown: number) => (total > shown ? ` (the first ${shown} shown)` : '')
/** A sort's direction as an arrow, in both sets. */
const sortArrow = (dir: 'asc' | 'desc') => (dir === 'asc' ? '↑' : '↓')

export const enMessages: SheetMessages = {
  back: 'Back to where you came from',
  backEmpty: 'No jumps yet',
  address: 'The selection’s address. Type one (B12, B2:D9, Sheet!C7) and press Enter to go there',
  addressBad: 'Not an address: try B12, B2:D9 or Sheet!C7',
  goSheet: (name) => `Go to sheet “${name}”`,
  showRange: 'Show the range',
  cellLink: 'Copy a link to this place in the book',
  cellLinkCopied: (ref) => `Link to ${ref} copied`,
  clipboardDenied: 'The browser denied access to the clipboard',
  freeze: (rows, cols, on) => {
    const what = [rows ? countEn(rows, 'row', 'rows') : '', cols ? countEn(cols, 'column', 'columns') : '']
      .filter(Boolean)
      .join(' and ')
    return on
      ? `The book freezes ${what} (the line marks the edge): click to unfreeze`
      : `Unfrozen, the line marks the edge: click to freeze ${what} as the book does`
  },
  note: (author) => (author ? `Note (${author})` : 'Note'),
  linkInBook: (location) => `In the book: ${location}`,
  linkOpen: 'Open',
  linkDead: (href) => `Link: ${href}`,
  loading: 'Reading the book…',
  empty: 'The sheet is empty',
  retry: 'Retry',
  loadFailed: 'Could not open the book',
  copyFailed: 'Could not copy the selection',
  copyPartial: 'Only part of the selection was copied: too many cells at once',
  statsCount: (n) => `Count: ${n}`,
  statsSum: (n) => `Sum: ${n}`,
  statsAverage: (n) => `Average: ${n}`,
  statsMin: (n) => `Min: ${n}`,
  statsMax: (n) => `Max: ${n}`,
  statsNumbers: (n) => `Numbers: ${n}`,
  statsCells: (n) => `Cells: ${n}`,
  sheetTitle: (name, rows, cols) => `${name}: ${countEn(rows, 'row', 'rows')}, ${countEn(cols, 'column', 'columns')}`,
  sheetUnnamed: (n) => `Sheet ${n}`,
  insertSheet: 'New sheet',
  sheetNamePrefix: 'Sheet',
  menu: 'Rows, columns and sheets',
  menuInsert: 'Insert',
  menuDelete: 'Delete',
  menuColumnWidth: 'Column Width…',
  menuRowHeight: 'Row Height…',
  menuHide: 'Hide',
  menuUnhide: 'Unhide',
  menuRename: 'Rename',
  menuMoveLeft: 'Move Left',
  menuMoveRight: 'Move Right',
  menuUnhideSheet: 'Unhide…',
  menuGroup: 'Group',
  menuUngroup: 'Ungroup',
  outlineRows: 'Rows',
  outlineColumns: 'Columns',
  outlineExpand: 'Show Detail',
  outlineCollapse: 'Hide Detail',
  outlineLevel: (n) => `Show level ${n}`,
  outlineDeepest: 'Groups go seven levels deep at most',
  outlineNone: 'There is no group here to ungroup',
  cellMenu: 'Cells',
  menuCut: 'Cut',
  menuCopy: 'Copy',
  menuPaste: 'Paste',
  pasteDenied: 'The browser does not let the page read the clipboard: paste with Ctrl+V (Cmd+V on a Mac)',
  fillHandle: 'Drag to fill the cells beside; double-click to fill down along the data',
  menuReapply: 'Reapply',
  filterReapply: 'Reapply the filter and sort to the data as they are now (Ctrl+Alt+L)',
  menuFreezePanes: 'Freeze Panes',
  menuUnfreezePanes: 'Unfreeze Panes',
  menuFreezeTopRow: 'Freeze Top Row',
  menuFreezeFirstColumn: 'Freeze First Column',
  menuPrecedents: 'Trace Precedents',
  explainTitle: 'Precedents',
  explainReading: 'Reading…',
  explainFailed: 'Could not read the cell’s precedents',
  explainEmpty: 'The source says nothing of this cell',
  columnWidth: 'Column width, in characters',
  rowHeight: 'Row height, in points',
  widthRange: 'Column width must be between 0 and 255',
  heightRange: 'Row height must be between 0 and 409',
  ok: 'OK',
  cancel: 'Cancel',
  deleteSheet: 'Delete the sheet',
  deleteSheetText: (name) => `The sheet “${name}” will be deleted with its data.`,
  unhideSheet: 'Unhide a sheet',
  rename: 'The sheet’s name',
  selectColumn: 'Click selects the column',
  selectRow: 'Click selects the row',
  sortColumn: 'Click sorts by this column',
  openRow: 'Open the row',
  selectAll: 'Click selects the whole sheet',
  resize: 'Drag to change the width; a double click restores the book’s width',
  resizeFit: 'Drag to change the width; a double click fits it to the contents',
  resizeRow: 'Drag to change the height; a double click fits it to the contents',
  conflict: 'The book was changed elsewhere: check and press Enter again',
  locked: (reason) => `This cell cannot be changed: ${reason}`,
  editFailed: 'The edit did not go through',
  pasteOffSheet: 'Cannot paste here: the data would go past the edge of the sheet',
  tooManyCells: 'Too many cells for one edit',
  pending: 'Saving…',
  close: 'Close',
  goRefHint: (key) => `${key}+click a reference to go to it`,
  barEnter: 'Enter',
  insertFunction: 'Insert Function',
  insertFunctionTitle: 'Insert Function',
  insertFunctionSearch: 'Search for a function',
  insertFunctionCategory: 'Or select a category',
  insertFunctionRecent: 'Most Recently Used',
  insertFunctionAll: 'All',
  insertFunctionPick: 'Select a function',
  expandBar: 'Expand Formula Bar (Ctrl+Shift+U)',
  collapseBar: 'Collapse Formula Bar (Ctrl+Shift+U)',
  waiting: 'The book is being recomputed…',
  waitSeconds: (n) => `${n} s`,
  waitHint: 'The values will show once the computation is done',
  waitTimeout: 'The book’s recomputation did not finish in time',
  find: 'Find on the sheet (Ctrl+F)',
  findPlaceholder: 'Find on the sheet…',
  findPrev: 'Previous (Shift+Enter)',
  findNext: 'Next (Enter)',
  findNone: 'no matches',
  findStatus: (i, total, shown) => `${i} of ${total}${shownEn(total, shown)}`,
  findCount: (total, shown) => `${total} found${shownEn(total, shown)}`,
  findFailed: 'The search failed',
  findOptions: 'Options',
  findMatchCase: 'Match case',
  findWholeCell: 'Match entire cell contents',
  findLookIn: 'Look in',
  findInFormulas: 'Formulas',
  findInValues: 'Values',
  findOrder: 'Search',
  findByRows: 'By Rows',
  findByColumns: 'By Columns',
  findWithin: (range) => `Within ${range}`,
  findWholeSheet: 'Search the whole sheet',
  replaceToggle: 'Replace (Ctrl+H)',
  replacePlaceholder: 'Replace with…',
  replaceOne: 'Replace',
  replaceAll: 'Replace All',
  replaceDone: (n) => (n ? `Replaced: ${n}` : 'Nothing to replace'),
  replaceRefused: (n, why) => `${countEn(n, 'cell was', 'cells were')} left as ${n === 1 ? 'it was' : 'they were'}: ${why}`,
  filter: 'Filter and sort',
  filterOn: 'A filter on this column is on',
  filterColumn: (letters) => `Column ${letters}`,
  filterNamed: (name, letters) => `${name} (${letters})`,
  sortAsc: 'Sort ascending',
  sortDesc: 'Sort descending',
  filterText: 'Contains text…',
  filterSearch: 'Find in the list…',
  filterAll: 'Select all',
  filterEmpty: '(empty)',
  filterMore: (n) => `${n} more values, narrow them with text`,
  filterReading: 'Reading the values…',
  filterApply: 'Apply',
  filterReset: 'Reset',
  filterClear: 'Clear the filter and sort',
  filterStatus: (shown, total, cols, sort) =>
    [`${shown} of ${total} rows shown`, cols.length ? `filter on ${cols.join(', ')}` : '', sort ? `sorted by ${sort.col} ${sortArrow(sort.dir)}` : '']
      .filter(Boolean)
      .join(', '),
  filterFailed: 'Could not apply the filter',
  valuesFailed: 'Could not read the column’s values',
  unfiltered: (row) => `The filter is cleared: it hid row ${row}`,
  toolbar: 'Formatting',
  toolbarMore: 'More formatting',
  undo: 'Undo (Ctrl+Z)',
  redo: 'Redo (Ctrl+Y)',
  bold: 'Bold',
  italic: 'Italic',
  underline: 'Underline',
  fontColour: 'Font colour',
  fillColour: 'Fill colour',
  pickColour: 'Pick a colour',
  colourAuto: 'Automatic',
  fillNone: 'No fill',
  borders: 'Borders',
  borderAll: 'All borders',
  borderOutside: 'Outside borders',
  borderBottom: 'Bottom border',
  borderNone: 'No border',
  alignLeft: 'Align left',
  alignCenter: 'Center',
  alignRight: 'Align right',
  wrapText: 'Wrap text',
  numberFormat: 'Number format',
  formatCells: 'Format Cells',
  formatCellsOpen: 'Format Cells… (Ctrl+1)',
  fmtTab: (tab) => ({ number: 'Number', alignment: 'Alignment', font: 'Font', border: 'Border', fill: 'Fill' })[tab],
  fmtCategory: 'Category',
  fmtCode: 'Type',
  fmtCodeEmpty: 'Type a format code',
  fmtHorizontal: 'Horizontal',
  fmtVertical: 'Vertical',
  fmtHorizontalKind: (h) =>
    ({ general: 'General', left: 'Left (Indent)', center: 'Center', right: 'Right (Indent)', fill: 'Fill', justify: 'Justify', centerContinuous: 'Center Across Selection', distributed: 'Distributed (Indent)' })[h],
  fmtVerticalKind: (v) => ({ top: 'Top', center: 'Center', bottom: 'Bottom', justify: 'Justify', distributed: 'Distributed' })[v],
  fmtIndent: 'Indent',
  fmtIndentRange: 'The indent must be a whole number from 0 to 250',
  fmtFontName: 'Font',
  fmtFontDefault: 'The book’s font',
  fmtFontStyle: 'Font style',
  fmtFontStyleKind: (k) => ({ regular: 'Regular', italic: 'Italic', bold: 'Bold', boldItalic: 'Bold Italic' })[k],
  fmtSize: 'Size',
  fmtSizeRange: 'The size must be between 1 and 409',
  fmtUnderline: 'Underline',
  fmtUnderlineKind: (k) => ({ none: 'None', single: 'Single', double: 'Double' })[k],
  fmtStrike: 'Strikethrough',
  fmtColour: 'Colour',
  fmtColourCode: 'Colour code, #RRGGBB',
  fmtColourBad: 'A colour is # and six hexadecimal digits',
  fmtBackground: 'Background colour',
  fmtLine: 'Line',
  fmtLineStyle: (style) =>
    ({
      hair: 'Hair',
      dotted: 'Dotted',
      dashDotDot: 'Dash dot dot',
      dashDot: 'Dash dot',
      dashed: 'Dashed',
      thin: 'Thin',
      mediumDashDotDot: 'Medium dash dot dot',
      slantDashDot: 'Slanted dash dot',
      mediumDashDot: 'Medium dash dot',
      mediumDashed: 'Medium dashed',
      medium: 'Medium',
      thick: 'Thick',
      double: 'Double',
    })[style] ?? style,
  fmtPresets: 'Presets',
  fmtPreset: (p) => ({ none: 'None', outline: 'Outline', inside: 'Inside' })[p],
  fmtEdges: 'Border',
  fmtEdge: (e) => ({ top: 'Top', bottom: 'Bottom', left: 'Left', right: 'Right', insideH: 'Inside horizontal', insideV: 'Inside vertical' })[e],
  validationList: 'Pick from the list (Alt+Down)',
  validationTitle: 'Microsoft Excel',
  validationYes: 'Yes',
  validationNo: 'No',
  validationContinue: 'Continue?',
  validationStyle: (severity) => ({ stop: 'Stop', warning: 'Warning', information: 'Information' })[severity],
  formulaQuestion: 'Fix the formula, or restore what the cell had?',
  formulaFix: 'Fix',
  formulaRevert: 'Restore',
  typoFound: (correction) => `We found a typo in your formula and tried to correct it to: ${correction}. Would you like to accept this correction?`,
  missingParenthesis: 'This formula is missing an opening or closing parenthesis.',
  functionList: 'Functions',
  functionHint: 'Tab or a double click enters the function',
  mergeCenter: 'Merge & Center',
  mergeMenu: 'Merge options',
  mergeAcross: 'Merge Across',
  mergeCells: 'Merge Cells',
  unmergeCells: 'Unmerge Cells',
  mergeWarning: 'Merging cells only keeps the upper-left value and discards other values.',
}

/** Russian plural: one, few, many. */
function plural(n: number, one: string, few: string, many: string): string {
  const m10 = n % 10
  const m100 = n % 100
  if (m10 === 1 && m100 !== 11) return one
  if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few
  return many
}
/** `n` and its Russian noun in the right plural. */
const countRu = (n: number, one: string, few: string, many: string) => `${n} ${plural(n, one, few, many)}`
/** The tail of a search status when the source cut the list. */
const shownRu = (total: number, shown: number) => (total > shown ? ` (показаны первые ${shown})` : '')

export const ruMessages: SheetMessages = {
  back: 'Вернуться туда, откуда перешли',
  backEmpty: 'Истории переходов пока нет',
  address: 'Адрес выделения. Введите адрес (B12, B2:D9, Лист!C7) и нажмите Enter, чтобы перейти',
  addressBad: 'Не понял адрес: нужно вроде B12, B2:D9 или Лист!C7',
  goSheet: (name) => `Перейти на лист «${name}»`,
  showRange: 'Показать диапазон',
  cellLink: 'Скопировать ссылку на это место в книге',
  cellLinkCopied: (ref) => `Ссылка на ${ref} скопирована`,
  clipboardDenied: 'Браузер не дал доступ к буферу обмена',
  freeze: (rows, cols, on) => {
    const what = [
      rows ? countRu(rows, 'строку', 'строки', 'строк') : '',
      cols ? countRu(cols, 'столбец', 'столбца', 'столбцов') : '',
    ]
      .filter(Boolean)
      .join(' и ')
    return on
      ? `В книге закреплены ${what} (граница отмечена линией): нажмите, чтобы снять закрепление`
      : `Закрепление снято, граница отмечена линией: нажмите, чтобы закрепить ${what} из книги`
  },
  note: (author) => (author ? `Примечание (${author})` : 'Примечание'),
  linkInBook: (location) => `В книге: ${location}`,
  linkOpen: 'Открыть',
  linkDead: (href) => `Ссылка: ${href}`,
  loading: 'Читаю книгу…',
  empty: 'Лист пуст',
  retry: 'Повторить',
  loadFailed: 'Не удалось открыть книгу',
  copyFailed: 'Не удалось скопировать выделение',
  copyPartial: 'Скопирована только часть выделения: слишком много ячеек за один раз',
  statsCount: (n) => `Ячеек: ${n}`,
  statsSum: (n) => `Сумма: ${n}`,
  statsAverage: (n) => `Среднее: ${n}`,
  statsMin: (n) => `Мин: ${n}`,
  statsMax: (n) => `Макс: ${n}`,
  statsNumbers: (n) => `Чисел: ${n}`,
  statsCells: (n) => `Выделено ячеек: ${n}`,
  sheetTitle: (name, rows, cols) => `${name}: ${countRu(rows, 'строка', 'строки', 'строк')}, ${countRu(cols, 'столбец', 'столбца', 'столбцов')}`,
  sheetUnnamed: (n) => `Лист ${n}`,
  insertSheet: 'Новый лист',
  sheetNamePrefix: 'Лист',
  menu: 'Строки, столбцы и листы',
  menuInsert: 'Вставить',
  menuDelete: 'Удалить',
  menuColumnWidth: 'Ширина столбца…',
  menuRowHeight: 'Высота строки…',
  menuHide: 'Скрыть',
  menuUnhide: 'Показать',
  menuRename: 'Переименовать',
  menuMoveLeft: 'Переместить влево',
  menuMoveRight: 'Переместить вправо',
  menuUnhideSheet: 'Показать…',
  menuGroup: 'Сгруппировать',
  menuUngroup: 'Разгруппировать',
  outlineRows: 'Строки',
  outlineColumns: 'Столбцы',
  outlineExpand: 'Показать детали',
  outlineCollapse: 'Скрыть детали',
  outlineLevel: (n) => `Показать уровень ${n}`,
  outlineDeepest: 'Группы вкладываются не глубже семи уровней',
  outlineNone: 'Здесь нет группы, разгруппировывать нечего',
  cellMenu: 'Ячейки',
  menuCut: 'Вырезать',
  menuCopy: 'Копировать',
  menuPaste: 'Вставить',
  pasteDenied: 'Браузер не даёт странице прочитать буфер обмена: вставьте сочетанием Ctrl+V (Cmd+V на Маке)',
  fillHandle: 'Потяните, чтобы заполнить соседние ячейки. Двойной щелчок заполняет вниз до конца данных',
  menuReapply: 'Применить повторно',
  filterReapply: 'Применить фильтр и сортировку повторно к нынешним данным (Ctrl+Alt+L)',
  menuFreezePanes: 'Закрепить области',
  menuUnfreezePanes: 'Снять закрепление областей',
  menuFreezeTopRow: 'Закрепить верхнюю строку',
  menuFreezeFirstColumn: 'Закрепить первый столбец',
  menuPrecedents: 'Влияющие ячейки',
  explainTitle: 'Влияющие ячейки',
  explainReading: 'Читаю…',
  explainFailed: 'Не удалось получить влияющие ячейки',
  explainEmpty: 'Источник ничего не сообщил об этой ячейке',
  columnWidth: 'Ширина столбца, в знаках',
  rowHeight: 'Высота строки, в пунктах',
  widthRange: 'Ширина столбца должна быть от 0 до 255',
  heightRange: 'Высота строки должна быть от 0 до 409',
  ok: 'ОК',
  cancel: 'Отмена',
  deleteSheet: 'Удалить лист',
  deleteSheetText: (name) => `Лист «${name}» будет удалён вместе с данными.`,
  unhideSheet: 'Показать лист',
  rename: 'Имя листа',
  selectColumn: 'Щелчок выделяет столбец',
  selectRow: 'Щелчок выделяет строку',
  sortColumn: 'Щелчок сортирует по этому столбцу',
  openRow: 'Открыть строку',
  selectAll: 'Щелчок выделяет весь лист',
  resize: 'Потяните, чтобы изменить ширину. Двойной щелчок вернёт ширину из книги',
  resizeFit: 'Потяните, чтобы изменить ширину. Двойной щелчок подберёт ширину по содержимому',
  resizeRow: 'Потяните, чтобы изменить высоту. Двойной щелчок подберёт высоту по содержимому',
  conflict: 'Книгу изменили в другом месте: проверьте и нажмите Enter ещё раз',
  locked: (reason) => `Эту ячейку менять нельзя: ${reason}`,
  editFailed: 'Правка не прошла',
  pasteOffSheet: 'Вставить нельзя: данные выйдут за край листа',
  tooManyCells: 'Слишком много ячеек для одной правки',
  pending: 'Сохраняю…',
  close: 'Закрыть',
  goRefHint: (key) => `${key}+щелчок по ссылке: перейти к ней`,
  barEnter: 'Ввод',
  insertFunction: 'Вставить функцию',
  insertFunctionTitle: 'Вставка функции',
  insertFunctionSearch: 'Поиск функции',
  insertFunctionCategory: 'Категория',
  insertFunctionRecent: '10 недавно использовавшихся',
  insertFunctionAll: 'Полный алфавитный перечень',
  insertFunctionPick: 'Выберите функцию',
  expandBar: 'Развернуть строку формул (Ctrl+Shift+U)',
  collapseBar: 'Свернуть строку формул (Ctrl+Shift+U)',
  waiting: 'Книга пересчитывается…',
  waitSeconds: (n) => `${n} с`,
  waitHint: 'Значения появятся, когда расчёт закончится',
  waitTimeout: 'Пересчёт книги не завершился за отведённое время',
  find: 'Поиск по листу (Ctrl+F)',
  findPlaceholder: 'Найти на листе…',
  findPrev: 'Предыдущее (Shift+Enter)',
  findNext: 'Следующее (Enter)',
  findNone: 'нет совпадений',
  findStatus: (i, total, shown) => `${i} из ${total}${shownRu(total, shown)}`,
  findCount: (total, shown) => `найдено: ${total}${shownRu(total, shown)}`,
  findFailed: 'Поиск не удался',
  findOptions: 'Параметры',
  findMatchCase: 'Учитывать регистр',
  findWholeCell: 'Ячейка целиком',
  findLookIn: 'Область поиска',
  findInFormulas: 'формулы',
  findInValues: 'значения',
  findOrder: 'Просматривать',
  findByRows: 'по строкам',
  findByColumns: 'по столбцам',
  findWithin: (range) => `В выделении ${range}`,
  findWholeSheet: 'Искать по всему листу',
  replaceToggle: 'Заменить (Ctrl+H)',
  replacePlaceholder: 'Заменить на…',
  replaceOne: 'Заменить',
  replaceAll: 'Заменить все',
  replaceDone: (n) => (n ? `Сделано замен: ${n}` : 'Нечего заменить'),
  replaceRefused: (n, why) => `Не заменено ячеек: ${n}. ${why}`,
  filter: 'Фильтр и сортировка',
  filterOn: 'Фильтр по этому столбцу включён',
  filterColumn: (letters) => `Столбец ${letters}`,
  filterNamed: (name, letters) => `${name} (${letters})`,
  sortAsc: 'Сортировать по возрастанию',
  sortDesc: 'Сортировать по убыванию',
  filterText: 'Содержит текст…',
  filterSearch: 'Найти в списке…',
  filterAll: 'Выбрать все',
  filterEmpty: '(пусто)',
  filterMore: (n) => `ещё ${n} значений, уточните подстрокой`,
  filterReading: 'Читаю значения…',
  filterApply: 'Применить',
  filterReset: 'Сбросить',
  filterClear: 'Снять фильтр и сортировку',
  filterStatus: (shown, total, cols, sort) =>
    [`показано ${shown} из ${total} строк`, cols.length ? `фильтр по ${cols.join(', ')}` : '', sort ? `сортировка по ${sort.col} ${sortArrow(sort.dir)}` : '']
      .filter(Boolean)
      .join(', '),
  filterFailed: 'Не удалось применить фильтр',
  valuesFailed: 'Не удалось получить значения столбца',
  unfiltered: (row) => `Фильтр снят: строка ${row} была скрыта им`,
  toolbar: 'Оформление',
  toolbarMore: 'Ещё оформление',
  undo: 'Отменить (Ctrl+Z)',
  redo: 'Вернуть (Ctrl+Y)',
  bold: 'Полужирный',
  italic: 'Курсив',
  underline: 'Подчёркнутый',
  fontColour: 'Цвет текста',
  fillColour: 'Цвет заливки',
  pickColour: 'Выбрать цвет',
  colourAuto: 'Авто',
  fillNone: 'Нет заливки',
  borders: 'Границы',
  borderAll: 'Все границы',
  borderOutside: 'Внешние границы',
  borderBottom: 'Нижняя граница',
  borderNone: 'Нет границы',
  alignLeft: 'По левому краю',
  alignCenter: 'По центру',
  alignRight: 'По правому краю',
  wrapText: 'Перенос текста',
  numberFormat: 'Числовой формат',
  formatCells: 'Формат ячеек',
  formatCellsOpen: 'Формат ячеек… (Ctrl+1)',
  fmtTab: (tab) => ({ number: 'Число', alignment: 'Выравнивание', font: 'Шрифт', border: 'Граница', fill: 'Заливка' })[tab],
  fmtCategory: 'Категория',
  fmtCode: 'Тип',
  fmtCodeEmpty: 'Введите код формата',
  fmtHorizontal: 'По горизонтали',
  fmtVertical: 'По вертикали',
  fmtHorizontalKind: (h) =>
    ({
      general: 'по значению',
      left: 'по левому краю (отступ)',
      center: 'по центру',
      right: 'по правому краю (отступ)',
      fill: 'с заполнением',
      justify: 'по ширине',
      centerContinuous: 'по центру выделения',
      distributed: 'распределённый (отступ)',
    })[h],
  fmtVerticalKind: (v) => ({ top: 'по верхнему краю', center: 'по центру', bottom: 'по нижнему краю', justify: 'по высоте', distributed: 'распределённый' })[v],
  fmtIndent: 'Отступ',
  fmtIndentRange: 'Отступ должен быть целым числом от 0 до 250',
  fmtFontName: 'Шрифт',
  fmtFontDefault: 'Шрифт книги',
  fmtFontStyle: 'Начертание',
  fmtFontStyleKind: (k) => ({ regular: 'обычный', italic: 'курсив', bold: 'полужирный', boldItalic: 'полужирный курсив' })[k],
  fmtSize: 'Размер',
  fmtSizeRange: 'Размер должен быть от 1 до 409',
  fmtUnderline: 'Подчёркивание',
  fmtUnderlineKind: (k) => ({ none: 'нет', single: 'одинарное', double: 'двойное' })[k],
  fmtStrike: 'Зачёркнутый',
  fmtColour: 'Цвет',
  fmtColourCode: 'Код цвета, #RRGGBB',
  fmtColourBad: 'Цвет записывается как # и шесть шестнадцатеричных цифр',
  fmtBackground: 'Цвет фона',
  fmtLine: 'Линия',
  fmtLineStyle: (style) =>
    ({
      hair: 'Волосяная',
      dotted: 'Точечная',
      dashDotDot: 'Штрихпунктирная с двумя точками',
      dashDot: 'Штрихпунктирная',
      dashed: 'Штриховая',
      thin: 'Тонкая',
      mediumDashDotDot: 'Средняя штрихпунктирная с двумя точками',
      slantDashDot: 'Наклонная штрихпунктирная',
      mediumDashDot: 'Средняя штрихпунктирная',
      mediumDashed: 'Средняя штриховая',
      medium: 'Средняя',
      thick: 'Толстая',
      double: 'Двойная',
    })[style] ?? style,
  fmtPresets: 'Все',
  fmtPreset: (p) => ({ none: 'нет', outline: 'внешние', inside: 'внутренние' })[p],
  fmtEdges: 'Отдельные',
  fmtEdge: (e) =>
    ({ top: 'верхняя', bottom: 'нижняя', left: 'левая', right: 'правая', insideH: 'внутренняя горизонтальная', insideV: 'внутренняя вертикальная' })[e],
  validationList: 'Выбрать из списка (Alt+↓)',
  validationTitle: 'Microsoft Excel',
  validationYes: 'Да',
  validationNo: 'Нет',
  validationContinue: 'Продолжить?',
  validationStyle: (severity) => ({ stop: 'Останов', warning: 'Предупреждение', information: 'Сообщение' })[severity],
  formulaQuestion: 'Исправить формулу или вернуть прежнее значение ячейки?',
  formulaFix: 'Исправить',
  formulaRevert: 'Вернуть',
  typoFound: (correction) => `Мы обнаружили опечатку в формуле и попытались ее исправить: ${correction}. Принять исправление?`,
  missingParenthesis: 'В формуле не хватает открывающей или закрывающей круглой скобки.',
  functionList: 'Функции',
  functionHint: 'Tab или двойной щелчок вставляет функцию',
  mergeCenter: 'Объединить и поместить в центре',
  mergeMenu: 'Параметры объединения',
  mergeAcross: 'Объединить по строкам',
  mergeCells: 'Объединить ячейки',
  unmergeCells: 'Отменить объединение ячеек',
  mergeWarning: 'При объединении ячеек сохраняется только значение из верхней левой ячейки, а значения из остальных ячеек удаляются.',
}

/** A component's texts: the English set under the part a host passed. */
export function useMessages(part: () => Partial<SheetMessages> | undefined): ComputedRef<SheetMessages> {
  return computed(() => ({ ...enMessages, ...part() }))
}
