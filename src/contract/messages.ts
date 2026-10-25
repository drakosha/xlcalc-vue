// The interface texts' keys and their arguments; the English and Russian sets are in src/messages.ts.

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
  /** The lock button's tooltip: what the workbook freezes, and whether the frozen panes are displayed now. */
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
  /** Selection totals over loaded cells: the count, then the sum, average, min, max and count of numbers when there are numbers. */
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
  /** The "+" button after the tabs, and the word a new sheet's name starts with (Excel's `Sheet4`, `Лист4`). */
  insertSheet: string
  sheetNamePrefix: string
  /** The structure menu on row numbers, column letters and sheet tabs, with Excel's item names. */
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
  /** Excel's outline: Group and Ungroup in the menu and the dialog for a cell range (rows or columns), the buttons' tooltips, and the messages for commands that are not performed. */
  menuGroup: string
  menuUngroup: string
  outlineRows: string
  outlineColumns: string
  outlineExpand: string
  outlineCollapse: string
  outlineLevel: (n: number) => string
  outlineDeepest: string
  outlineNone: string
  /** Excel's cell context menu (cut, copy and paste within the workbook), a paste whose clipboard the browser does not allow to read, and the fill handle. */
  cellMenu: string
  menuCut: string
  menuCopy: string
  menuPaste: string
  pasteDenied: string
  fillHandle: string
  /** Excel's commands in the cell context menu: Reapply for the display filter, Freeze Panes and its variants, Trace Precedents as a panel next to the grid. */
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
  /** The size dialogs: their labels in Excel's units, and the limits. */
  columnWidth: string
  rowHeight: string
  widthRange: string
  heightRange: string
  ok: string
  cancel: string
  /** The confirmation before a sheet with data is deleted, and the list of hidden sheets. */
  deleteSheet: string
  deleteSheetText: (name: string) => string
  unhideSheet: string
  rename: string
  selectColumn: string
  selectRow: string
  /** List mode: a sortable column's title, and the cells of the first column that open their row. */
  sortColumn: string
  openRow: string
  selectAll: string
  resize: string
  /** A column edge where a double click fits the column to its contents. */
  resizeFit: string
  /** The resize grip on a row number's edge. */
  resizeRow: string
  conflict: string
  locked: (reason: string) => string
  editFailed: string
  /** A paste whose range would extend past the sheet's last row or column. */
  pasteOffSheet: string
  tooManyCells: string
  pending: string
  close: string
  /** The formula bar's tooltip over a formula with references: Ctrl+click (Cmd on a Mac; the key is passed in) goes to a reference. */
  goRefHint: (key: string) => string
  /** Excel's Enter button to the left of the formula bar (its Cancel button is `cancel`): enters the edit and keeps the selection. */
  barEnter: string
  /** Excel's Insert Function: the fx button, then the dialog's title, search box, category box with its first two items, and the list's label. */
  insertFunction: string
  insertFunctionTitle: string
  insertFunctionSearch: string
  insertFunctionCategory: string
  insertFunctionRecent: string
  insertFunctionAll: string
  insertFunctionPick: string
  /** Excel's button at the right end of the formula bar and its Ctrl+Shift+U: expands or collapses the formula bar. */
  expandBar: string
  collapseBar: string
  /** A workbook being recalculated: the waiting message, the seconds elapsed, and the timeout. */
  waiting: string
  waitSeconds: (n: number) => string
  waitHint: string
  waitTimeout: string
  find: string
  findPlaceholder: string
  findPrev: string
  findNext: string
  findNone: string
  /** Search status: the current match's position, the total, and how many were received when the source truncated the list. */
  findStatus: (i: number, total: number, shown: number) => string
  /** Search status while the active cell is not a match: the total, and how many were received. */
  findCount: (total: number, shown: number) => string
  findFailed: string
  /** Excel's Find and Replace: its options, the range it searches within, the Replace row and the result of Replace All. */
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
  /** Replace All left `n` cells unchanged: Excel would not accept their new text, for the source's reason `why`. */
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
  /** The filter status line at the bottom: visible rows out of all, the columns with conditions, the sort. */
  filterStatus: (shown: string, total: string, cols: string[], sort: { col: string; dir: 'asc' | 'desc' } | null) => string
  filterFailed: string
  valuesFailed: string
  unfiltered: (row: number) => string
  /** The formatting toolbar (`SheetToolbar`): its buttons, menus and number formats, with Excel's names. */
  toolbar: string
  toolbarMore: string
  undo: string
  redo: string
  bold: string
  italic: string
  underline: string
  fontColor: string
  fillColor: string
  pickColor: string
  colorAuto: string
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
  /** Excel's Format Cells dialog (`FormatDialog`): its tabs, inputs and the names of their options. */
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
  fmtColor: string
  fmtColorCode: string
  fmtColorBad: string
  fmtBackground: string
  fmtLine: string
  fmtLineStyle: (style: string) => string
  fmtPresets: string
  fmtPreset: (p: 'none' | 'outline' | 'inside') => string
  fmtEdges: string
  fmtEdge: (e: 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV') => string
  /** Excel's data validation: the dropdown button, the alert's title when the rule has none, its buttons (Retry, OK and Cancel are `retry`, `ok` and `cancel`), a warning's question and the style names. */
  validationList: string
  validationTitle: string
  validationYes: string
  validationNo: string
  validationContinue: string
  validationStyle: (severity: 'stop' | 'warning' | 'information') => string
  /** The prompt for a formula that cannot be parsed, below the source's reason, and its buttons. */
  formulaQuestion: string
  formulaFix: string
  formulaRevert: string
  /** Excel's dialog for an unclosed formula that it offers to correct (Yes and No are `validationYes` and `validationNo`), and its message after No. */
  typoFound: (correction: string) => string
  missingParenthesis: string
  /** Excel's Formula AutoComplete: the function list below the editor and its hint. */
  functionList: string
  functionHint: string
  /** Excel's Merge & Center button, its dropdown button and items, and the warning before values are lost. */
  mergeCenter: string
  mergeMenu: string
  mergeAcross: string
  mergeCells: string
  unmergeCells: string
  mergeWarning: string
  /** Excel's Trace Dependents: the message when no formula refers to the active cell, and when the source could not be asked. */
  traceNone: string
  traceFailed: string
  /** Excel's Create Table dialog, the range it refuses, the words a new table's and an unnamed column's names start with (`Table1`, `Column1`), and Convert to Range's question. */
  tableCreate: string
  tableWhere: string
  tableHeaders: string
  tableRangeBad: string
  tableNamePrefix: string
  tableColumnPrefix: string
  tableConvert: string
  /** Excel's Spelling dialog: its title, boxes and buttons (Cancel is `cancel`), the message when the check is complete, and the one when it fails. */
  spelling: string
  spellNotInDictionary: string
  spellSuggestions: string
  spellNoSuggestions: string
  spellIgnoreOnce: string
  spellIgnoreAll: string
  spellAdd: string
  spellChange: string
  spellChangeAll: string
  spellDone: string
  spellFailed: string
}
