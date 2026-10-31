// The interface texts' keys and their arguments; the English and Russian sets are in src/messages.ts.
import type { ConditionalFormat, IconSetName, SheetProtection, TimelineLevel } from 'xlcalc/contract'
import type { QuickRule } from './conditional'

export interface SheetMessages extends ConditionalMessages, ProtectionMessages, PivotFieldsMessages, SlicerMessages {
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
  /** Excel's Tab Color on the tab menu, and its No Color. */
  menuTabColor: string
  tabColorNone: string
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
  /** Excel's Paste Special: the cell menu's item, the dialog's title, its Paste group and choices, its Operation group, Skip Blanks and Transpose. */
  menuPasteSpecial: string
  pasteSpecialTitle: string
  pasteSpecialPaste: string
  pasteSpecialAll: string
  pasteSpecialFormulas: string
  pasteSpecialValues: string
  pasteSpecialFormats: string
  pasteSpecialValidation: string
  pasteSpecialTranspose: string
  pasteSpecialNotes: string
  pasteSpecialAllExceptBorders: string
  pasteSpecialColumnWidths: string
  pasteSpecialFormulasNumbers: string
  pasteSpecialValuesNumbers: string
  pasteSpecialMergeConditional: string
  pasteSpecialOperation: string
  pasteSpecialNone: string
  pasteSpecialAdd: string
  pasteSpecialSubtract: string
  pasteSpecialMultiply: string
  pasteSpecialDivide: string
  pasteSpecialSkipBlanks: string
  /** Excel's Format Painter on the toolbar. */
  formatPainter: string
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
  /** Excel's refusal of a copy, a cut or a paste over a selection of several areas (Go To Special's). */
  multipleSelections: string
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
  /** Data > Filter with no data at the active cell or in the selection, Excel's message. */
  autoFilterNoRange: string
  /** The formatting toolbar (`SheetToolbar`): its buttons, menus and number formats, with Excel's names. */
  toolbar: string
  toolbarMore: string
  undo: string
  redo: string
  bold: string
  italic: string
  underline: string
  fontColor: string
  /** The Font and Font Size boxes, and Increase and Decrease Font Size. */
  fontName: string
  fontSize: string
  fontSizeUp: string
  fontSizeDown: string
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
  alignTop: string
  alignMiddle: string
  alignBottom: string
  indentDecrease: string
  indentIncrease: string
  wrapText: string
  numberFormat: string
  /** Excel's Percent Style button, with its key. */
  percentStyle: string
  /** Excel's Accounting Number Format, Comma Style, Increase Decimal and Decrease Decimal buttons. */
  accountingStyle: string
  commaStyle: string
  decimalIncrease: string
  decimalDecrease: string
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
  /** The Alignment tab's Orientation: its box, the degrees, Vertical Text, and the degrees' limits. */
  fmtOrientation: string
  fmtDegrees: string
  fmtVerticalText: string
  fmtDegreesRange: string
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
  /** Excel's AutoSum on the toolbar: the button (Sum), its dropdown button, and the list's items. */
  autoSum: string
  autoSumMenu: string
  autoSumSum: string
  autoSumAverage: string
  autoSumCount: string
  autoSumMax: string
  autoSumMin: string
  autoSumMore: string
  /** Excel's Go To dialog (F5, Ctrl+G): its title, the list's label, the Reference box, and its message for text that is not a reference. */
  goToTitle: string
  goToList: string
  goToReference: string
  goToBad: string
  /** Excel's Go To Special: the Go To dialog's Special… button, the dialog's title, its Select group and options, and Excel's message when it finds no cells. */
  goToSpecial: string
  specialTitle: string
  specialSelect: string
  specialConstants: string
  specialFormulas: string
  specialNumbers: string
  specialText: string
  specialLogicals: string
  specialErrors: string
  specialBlanks: string
  specialRegion: string
  specialLastCell: string
  specialVisible: string
  specialConditional: string
  specialValidation: string
  specialAll: string
  specialSame: string
  specialNone: string
  /** Excel's Symbol dialog: its title and tabs, the Subset box and its blocks, the recently used row, the Character code box, the Special Characters tab's names, Insert. */
  symbolTitle: string
  symbolTab: (tab: 'symbols' | 'special') => string
  symbolSubsetLabel: string
  symbolSubset: (subset: SymbolSubset) => string
  symbolRecent: string
  symbolCode: string
  symbolCodeFrom: string
  symbolCharacter: string
  symbolSpecial: (character: SymbolSpecial) => string
  symbolInsert: string
  /** Defined names: the Name Box's list button, its messages for a typed name, the cell menu's Define Name, and a name in Formula AutoComplete. */
  nameBoxList: string
  nameBoxBad: string
  nameNotRange: string
  menuDefineName: string
  nameItem: (name: string) => string
  /** Excel's New Name and Edit Name dialog and its fields; the scope that is not a sheet. */
  nameNewTitle: string
  nameEditTitle: string
  nameLabel: string
  nameScope: string
  nameRefersTo: string
  nameWorkbook: string
  /** Excel's Name Manager: its title, buttons, the list's columns, Commit beside Refers to, and its questions before a deletion and before closing with an edited reference. */
  nameManager: string
  nameNew: string
  nameEdit: string
  nameDelete: string
  nameColumn: (column: 'name' | 'refersTo' | 'scope') => string
  nameCommit: string
  nameDeleteAsk: (name: string) => string
  nameSaveAsk: string
  /** Excel's Create Names from Selection: the dialog's title, its caption and boxes, Excel's message for labels alone, and its question before a name's other definition is replaced. */
  createNamesTitle: string
  createNamesIn: string
  createNamesTop: string
  createNamesLeft: string
  createNamesBottom: string
  createNamesRight: string
  createNamesInvalid: string
  createNamesReplace: (name: string) => string
  /** Excel's Workbook Statistics: the window's title, its two sections and the labels of its rows. */
  statisticsTitle: string
  statisticsSheet: string
  statisticsWorkbook: string
  statisticsEnd: string
  statisticsSheets: string
  statisticsCells: string
  statisticsTables: string
  statisticsPivotTables: string
  statisticsFormulas: string
  statisticsCharts: string
  statisticsImages: string
  statisticsFormControls: string
  statisticsObjects: string
  statisticsComments: string
  statisticsNotes: string
  /** Excel's zoom on the status bar (Zoom Out, the slider, Zoom In, the percentage that opens the Zoom dialog), and the dialog's title, choices and limits. */
  zoomOut: string
  zoomIn: string
  zoomSlider: string
  zoomLevel: string
  zoomTitle: string
  zoomFit: string
  zoomCustom: string
  zoomRange: string
  /** Excel's Watch Window: its title, Add Watch and Delete Watch, its columns, and the cap on the cells watched. */
  watchTitle: string
  watchAdd: string
  watchDelete: string
  watchBook: string
  watchSheet: string
  watchName: string
  watchCell: string
  watchValue: string
  watchFormula: string
  watchTooMany: (max: number) => string
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
  /** Excel's hyperlink items in the cell menu and its Insert Hyperlink dialog (Edit Hyperlink on a cell with a link), with Excel's names. */
  menuLink: string
  menuEditLink: string
  menuOpenLink: string
  menuRemoveLink: string
  linkInsertTitle: string
  linkEditTitle: string
  linkTo: string
  linkToFile: string
  linkToPlace: string
  linkText: string
  /** Text to Display over several cells or a formula, which the link does not change. */
  linkSelection: string
  linkTip: string
  linkAddress: string
  linkRef: string
  linkPlaces: string
  linkCellReference: string
  linkDefinedNames: string
  linkRemove: string
  linkBadRef: string
  /** Excel's note items in the cell menu (New Note, Edit Note, Delete Note) and the note box's label. */
  menuNewNote: string
  menuEditNote: string
  menuDeleteNote: string
  /** Excel's Show/Hide Note on a cell with a note. */
  menuShowNote: string
  noteEditor: string
  /** Excel's Data Validation dialog: its title and tabs, the Settings tab's boxes by the rule's type and comparison, the Input Message and Error Alert tabs (the styles are `validationStyle`), Clear All, Excel's questions before it opens and its checks before OK. */
  dataValidation: string
  dvTab: (tab: 'settings' | 'input' | 'error') => string
  dvCriteria: string
  dvAllow: string
  dvType: (type: 'any' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom') => string
  dvData: string
  dvOperator: (op: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual') => string
  dvBox: (box: 'min' | 'max' | 'value' | 'length' | 'start' | 'end' | 'date' | 'startTime' | 'endTime' | 'time' | 'source' | 'formula') => string
  dvIgnoreBlank: string
  dvDropdown: string
  dvApplyAll: string
  dvShowPrompt: string
  dvPromptLead: string
  dvFieldTitle: string
  dvPromptText: string
  dvShowError: string
  dvErrorLead: string
  dvStyle: string
  dvErrorText: string
  dvClearAll: string
  dvMany: string
  dvExtend: string
  /** OK refused: an empty box, a box whose value is not one, a maximum below the minimum; with the boxes' names from `dvBox`. */
  dvEmpty: (box: string) => string
  dvInvalid: (box: string) => string
  dvOrder: (min: string, max: string) => string
  /** The status bar's Calculate in Manual while cells wait for a calculation, and its tooltip (a click calculates, as F9). */
  calculateStatus: string
  calculateHint: string
  /** Excel's Data > Refresh All on the toolbar, with its key; its list's button and items, Refresh All and Refresh (of the PivotTable under the active cell) with its key. */
  refreshAll: string
  refreshAllMenu: string
  refreshAllItem: string
  refreshPivot: string
  /** Refresh in the cell menu of a PivotTable's cell. */
  menuRefresh: string
}

/** The Symbol dialog's Subset list: Unicode blocks by Excel's names (`basicLatin` is Basic Latin, `latin1` Latin-1 Supplement). */
export type SymbolSubset =
  | 'basicLatin'
  | 'latin1'
  | 'latinExtA'
  | 'latinExtB'
  | 'ipa'
  | 'spacingModifiers'
  | 'combiningMarks'
  | 'greek'
  | 'cyrillic'
  | 'cyrillicSupplement'
  | 'phoneticExt'
  | 'latinExtAdditional'
  | 'greekExt'
  | 'punctuation'
  | 'superSub'
  | 'currency'
  | 'letterlike'
  | 'numberForms'
  | 'arrows'
  | 'math'
  | 'technical'
  | 'enclosed'
  | 'boxDrawing'
  | 'blocks'
  | 'geometric'
  | 'miscSymbols'
  | 'dingbats'
  | 'latinExtC'
  | 'presentationForms'
  | 'specials'

/** The characters of the Symbol dialog's Special Characters tab, in Excel's order. */
export type SymbolSpecial =
  | 'emDash'
  | 'enDash'
  | 'nonBreakingHyphen'
  | 'optionalHyphen'
  | 'emSpace'
  | 'enSpace'
  | 'quarterEmSpace'
  | 'nonBreakingSpace'
  | 'copyright'
  | 'registered'
  | 'trademark'
  | 'section'
  | 'paragraph'
  | 'ellipsis'
  | 'singleOpeningQuote'
  | 'singleClosingQuote'
  | 'doubleOpeningQuote'
  | 'doubleClosingQuote'
  | 'noWidthOptionalBreak'
  | 'noWidthNonBreak'
type CfOperator = Extract<ConditionalFormat, { type: 'cellIs' }>['operator']
type CfTextType = 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'
type CfPeriod = Extract<ConditionalFormat, { type: 'timePeriod' }>['period']
type CfValueType = 'min' | 'max' | 'num' | 'percent' | 'formula' | 'percentile' | 'autoMin' | 'autoMax'

/** Excel's conditional formatting dialogs: the quick rules' dialogs, the Rules Manager, New and Edit Formatting Rule, a rule's Format Cells. */
export interface ConditionalMessages {
  /** A quick rule's dialog: its title, the label above its box, and the "with" box's formats, Custom Format last. */
  cfQuickTitle: (kind: QuickRule) => string
  cfQuickLabel: (kind: QuickRule) => string
  cfWith: string
  cfAnd: string
  cfValuesWith: string
  cfForRange: string
  cfPreset: (preset: 'lightRedDarkRed' | 'yellowDarkYellow' | 'greenDarkGreen' | 'lightRed' | 'redText' | 'redBorder') => string
  cfCustomFormat: string
  cfPeriod: (period: CfPeriod) => string
  cfDuplicate: (unique: boolean) => string
  /** Why OK does not write the rule. */
  cfValueMissing: string
  cfRankRange: (max: number) => string
  cfPercentRange: string
  cfRelative: string
  cfAppliesBad: string
  /** The Rules Manager: its list's choice, buttons, columns and the rules' descriptions. */
  cfManager: string
  cfShowFor: string
  cfShowSelection: string
  cfShowSheet: string
  cfNewRule: string
  cfEditRule: string
  cfDeleteRule: string
  cfDuplicateRule: string
  cfMoveUp: string
  cfMoveDown: string
  cfColRule: string
  cfColFormat: string
  cfColApplies: string
  cfColStop: string
  cfApply: string
  cfLoading: string
  cfNoFormat: string
  cfSample: string
  cfDescCell: (operator: CfOperator, a: string, b?: string) => string
  cfDescFormula: (formula: string) => string
  cfDescText: (type: CfTextType, text: string) => string
  cfDescBlanks: (type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors') => string
  cfDescTop: (rank: number, percent: boolean, bottom: boolean) => string
  cfDescAverage: (below: boolean, equal: boolean, stdDev?: number) => string
  cfDescDuplicate: (unique: boolean) => string
  cfDescScale: string
  cfDescBar: string
  cfDescIcons: string
  /** New Formatting Rule and Edit Formatting Rule: the rule types and each one's inputs. */
  cfNewTitle: string
  cfEditTitle: string
  cfSelectType: string
  cfType: (kind: 'values' | 'contain' | 'rank' | 'average' | 'unique' | 'formula') => string
  cfEditDescription: string
  cfFormatStyle: string
  cfStyleKind: (style: 'scale2' | 'scale3' | 'dataBar' | 'iconSet') => string
  cfPoint: (which: 'min' | 'mid' | 'max') => string
  cfTypeLabel: string
  cfValueLabel: string
  cfColorLabel: string
  cfValueType: (type: CfValueType) => string
  cfBarAppearance: string
  cfBarFill: (gradient: boolean) => string
  cfBarBorder: (on: boolean) => string
  cfShowBarOnly: string
  cfDirection: string
  cfDirectionKind: (d: 'context' | 'leftToRight' | 'rightToLeft') => string
  cfIconStyle: string
  cfIconSet: (set: IconSetName) => string
  cfReverse: string
  cfShowIconOnly: string
  cfIconRules: string
  cfWhenValue: string
  cfIconLast: string
  cfFormatOnly: string
  cfContain: (kind: 'cellValue' | 'text' | 'dates' | 'blanks' | 'noBlanks' | 'errors' | 'noErrors') => string
  cfOperator: (operator: CfOperator) => string
  cfTextOp: (type: CfTextType) => string
  cfRankLabel: string
  cfTopBottom: (top: boolean) => string
  cfPercentOf: string
  cfAverageLabel: string
  cfAverageKind: (kind: 'above' | 'below' | 'equalAbove' | 'equalBelow' | 'std1Above' | 'std2Above' | 'std3Above' | 'std1Below' | 'std2Below' | 'std3Below') => string
  cfAverageOf: string
  cfUniqueLabel: string
  cfInRange: string
  cfFormulaLabel: string
  cfPreview: string
  cfFormatButton: string
  /** A rule's Format Cells: no choice in a list, and Clear. */
  cfNotSet: string
  cfClear: string
}

/** The texts of sheet protection, in Excel for Windows' words: its alert, the tab menu's commands, the Protect Sheet dialog, Format Cells' Protection tab. */
export interface ProtectionMessages {
  /** Excel's alert for a change a protected sheet does not allow, such as typing into a locked cell. */
  protectedAlert: string
  /** The sheet tab's menu: Protect Sheet… and Unprotect Sheet. */
  menuProtectSheet: string
  menuUnprotectSheet: string
  /** The Protect Sheet dialog: its title, the box that protects the sheet, the list's heading, and each of its boxes. */
  protectTitle: string
  protectContents: string
  protectAllow: string
  protectBox: (box: keyof SheetProtection) => string
  /** Format Cells' Protection tab: its name, Locked, Hidden, and Excel's note. */
  fmtProtection: string
  fmtLocked: string
  fmtHidden: string
  fmtProtectionNote: string
}

/** The texts of Insert PivotTable and the PivotTable Fields pane, in Excel for Windows' words. */
export interface PivotFieldsMessages {
  /** Insert > PivotTable on the toolbar. */
  pivotInsert: string
  /** The PivotTable from table or range dialog: its title, the source's heading and box, the place's heading, its choices and box. */
  pivotCreateTitle: string
  pivotCreateSource: string
  pivotCreateRange: string
  pivotCreatePlace: string
  pivotNewSheet: string
  pivotExistingSheet: string
  pivotLocation: string
  /** Excel's alerts for a source or a location the dialog does not take. */
  pivotSourceBad: string
  pivotLocationBad: string
  /** The pane: its title, the heading over the fields, the search box, the heading over the areas, and the areas. */
  fieldListTitle: string
  fieldListChoose: string
  fieldListSearch: string
  fieldListDrag: string
  areaFilters: string
  areaColumns: string
  areaRows: string
  areaValues: string
  /** The value fields' place on Columns or Rows when there are several (Σ Values). */
  fieldListValues: string
  /** The arrow beside a field in the list, which opens its sort and filter menu. */
  fieldListFilter: string
  /** Defer Layout Update and its Update button. */
  fieldListDefer: string
  fieldListUpdate: string
  /** A field's menu in an area. */
  fieldMoveUp: string
  fieldMoveDown: string
  fieldMoveFirst: string
  fieldMoveLast: string
  fieldToFilters: string
  fieldToRows: string
  fieldToColumns: string
  fieldToValues: string
  fieldRemove: string
  fieldSettings: string
  valueFieldSettings: string
  /** A field's menu in the list. */
  fieldAddFilters: string
  fieldAddRows: string
  fieldAddColumns: string
  fieldAddValues: string
  /** The cell menu of a PivotTable's cell: Show Field List and Hide Field List. */
  menuShowFieldList: string
  menuHideFieldList: string
}

/**
 * The texts of slicers and timelines in Excel for Windows' words: the slicer's header buttons with their keys, the
 * timeline's levels, periods and selection label, the menu of a slicer or a timeline, and the dialogs Insert Slicers,
 * Insert Timelines, Slicer Settings and Report Connections.
 */
export interface SlicerMessages {
  /** The slicer's header buttons with their keys, Clear Filter also the timeline's. */
  slicerMultiSelect: string
  slicerClearFilter: string
  /** The timeline's label for no dates chosen, and the level picker's levels. */
  timelineAllPeriods: string
  timelineLevel: (level: TimelineLevel) => string
  /** A quarter's tile (1 to 4), and the twelve months' short names, for the tiles and the selection label. */
  timelineQuarter: (n: number) => string
  timelineMonths: string[]
  /** The selection label: one day; one period of a year; whole years; periods within a year; periods across years. */
  timelineDay: (day: number, month: string, year: number) => string
  timelinePeriod: (period: string, year: number) => string
  timelineYears: (from: string, to: string) => string
  timelineInYear: (from: string, to: string, year: number) => string
  timelineAcross: (from: string, fromYear: number, to: string, toYear: number) => string
  /** The menu of a slicer or a timeline: sorting, Clear Filter from, Multi-Select and Remove with its name, the dialogs. */
  menuSortAZ: string
  menuSortZA: string
  menuClearFilterFrom: (name: string) => string
  menuMultiSelect: (name: string) => string
  menuRemoveObject: (name: string) => string
  menuReportConnections: string
  menuSlicerSettings: string
  /** The dialogs Insert Slicers and Insert Timelines, and Excel's alert over a PivotTable without dates. */
  insertSlicersTitle: string
  insertTimelinesTitle: string
  timelineNoDates: string
  /** Slicer Settings: the title, the names, the Header group, Item Sorting and Filtering with its options. */
  slicerSettingsTitle: string
  slicerSourceName: string
  slicerFormulaName: string
  slicerName: string
  slicerHeader: string
  slicerDisplayHeader: string
  slicerCaption: string
  slicerSorting: string
  slicerAscending: string
  slicerDescending: string
  slicerCustomLists: string
  slicerHideNoData: string
  slicerMarkNoData: string
  slicerNoDataLast: string
  slicerShowMissing: string
  /** Report Connections: the title with the slicer's or timeline's name, the line above the list, and the list's columns. */
  reportConnectionsTitle: (name: string) => string
  reportConnectionsPrompt: string
  reportConnectionsName: string
  reportConnectionsSheet: string
}
