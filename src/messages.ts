// Interface texts of the component: English by default, a Russian set in the package; a host passes its own or a part.

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
  selectAll: string
  resize: string
  /** A column's edge where a double click fits the column to what it shows. */
  resizeFit: string
  conflict: string
  locked: (reason: string) => string
  editFailed: string
  /** A paste whose rectangle would go past the sheet's last row or column. */
  pasteOffSheet: string
  tooManyCells: string
  pending: string
  close: string
  editHint: string
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
  formatGeneral: string
  formatNumber: string
  formatCurrency: string
  formatPercent: string
  formatDate: string
  formatTime: string
  formatText: string
  /** A format none of the offered ones is (Excel's Custom). */
  formatOther: string
}

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
    const what = [rows ? `${rows} ${rows === 1 ? 'row' : 'rows'}` : '', cols ? `${cols} ${cols === 1 ? 'column' : 'columns'}` : '']
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
  sheetTitle: (name, rows, cols) => `${name}: ${rows} ${rows === 1 ? 'row' : 'rows'}, ${cols} ${cols === 1 ? 'column' : 'columns'}`,
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
  selectAll: 'Click selects the whole sheet',
  resize: 'Drag to change the width; a double click restores the book’s width',
  resizeFit: 'Drag to change the width; a double click fits it to the contents',
  conflict: 'The book was changed elsewhere: check and press Enter again',
  locked: (reason) => `This cell cannot be changed: ${reason}`,
  editFailed: 'The edit did not go through',
  pasteOffSheet: 'Cannot paste here: the data would go past the edge of the sheet',
  tooManyCells: 'Too many cells for one edit',
  pending: 'Saving…',
  close: 'Close',
  editHint: 'Click to edit the cell',
  waiting: 'The book is being recomputed…',
  waitSeconds: (n) => `${n} s`,
  waitHint: 'The values will show once the computation is done',
  waitTimeout: 'The book’s recomputation did not finish in time',
  find: 'Find on the sheet (Ctrl+F)',
  findPlaceholder: 'Find on the sheet…',
  findPrev: 'Previous (Shift+Enter)',
  findNext: 'Next (Enter)',
  findNone: 'no matches',
  findStatus: (i, total, shown) => `${i} of ${total}${total > shown ? ` (the first ${shown} shown)` : ''}`,
  findCount: (total, shown) => `${total} found${total > shown ? ` (the first ${shown} shown)` : ''}`,
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
    [`${shown} of ${total} rows shown`, cols.length ? `filter on ${cols.join(', ')}` : '', sort ? `sorted by ${sort.col} ${sort.dir === 'asc' ? '↑' : '↓'}` : '']
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
  formatGeneral: 'General',
  formatNumber: 'Number',
  formatCurrency: 'Currency',
  formatPercent: 'Percentage',
  formatDate: 'Short date',
  formatTime: 'Time',
  formatText: 'Text',
  formatOther: 'Custom',
}

/** Russian plural: one, few, many. */
function plural(n: number, one: string, few: string, many: string): string {
  const m10 = n % 10
  const m100 = n % 100
  if (m10 === 1 && m100 !== 11) return one
  if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few
  return many
}

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
      rows ? `${rows} ${plural(rows, 'строку', 'строки', 'строк')}` : '',
      cols ? `${cols} ${plural(cols, 'колонку', 'колонки', 'колонок')}` : '',
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
  sheetTitle: (name, rows, cols) =>
    `${name}: ${rows} ${plural(rows, 'строка', 'строки', 'строк')}, ${cols} ${plural(cols, 'колонка', 'колонки', 'колонок')}`,
  sheetUnnamed: (n) => `Лист ${n}`,
  insertSheet: 'Новый лист',
  sheetNamePrefix: 'Лист',
  menu: 'Строки, колонки и листы',
  menuInsert: 'Вставить',
  menuDelete: 'Удалить',
  menuColumnWidth: 'Ширина колонки…',
  menuRowHeight: 'Высота строки…',
  menuHide: 'Скрыть',
  menuUnhide: 'Показать',
  menuRename: 'Переименовать',
  menuMoveLeft: 'Переместить влево',
  menuMoveRight: 'Переместить вправо',
  menuUnhideSheet: 'Показать…',
  columnWidth: 'Ширина колонки, в знаках',
  rowHeight: 'Высота строки, в пунктах',
  widthRange: 'Ширина колонки должна быть от 0 до 255',
  heightRange: 'Высота строки должна быть от 0 до 409',
  ok: 'ОК',
  cancel: 'Отмена',
  deleteSheet: 'Удалить лист',
  deleteSheetText: (name) => `Лист «${name}» будет удалён вместе с данными.`,
  unhideSheet: 'Показать лист',
  rename: 'Имя листа',
  selectColumn: 'Щелчок выделяет колонку',
  selectRow: 'Щелчок выделяет строку',
  selectAll: 'Щелчок выделяет весь лист',
  resize: 'Потяните, чтобы изменить ширину. Двойной щелчок вернёт ширину из книги',
  resizeFit: 'Потяните, чтобы изменить ширину. Двойной щелчок подберёт ширину по содержимому',
  conflict: 'Книгу изменили в другом месте: проверьте и нажмите Enter ещё раз',
  locked: (reason) => `Эту ячейку менять нельзя: ${reason}`,
  editFailed: 'Правка не прошла',
  pasteOffSheet: 'Вставить нельзя: данные выйдут за край листа',
  tooManyCells: 'Слишком много ячеек для одной правки',
  pending: 'Сохраняю…',
  close: 'Закрыть',
  editHint: 'Нажмите, чтобы править ячейку',
  waiting: 'Книга пересчитывается…',
  waitSeconds: (n) => `${n} с`,
  waitHint: 'Значения появятся, когда расчёт закончится',
  waitTimeout: 'Пересчёт книги не завершился за отведённое время',
  find: 'Поиск по листу (Ctrl+F)',
  findPlaceholder: 'Найти на листе…',
  findPrev: 'Предыдущее (Shift+Enter)',
  findNext: 'Следующее (Enter)',
  findNone: 'нет совпадений',
  findStatus: (i, total, shown) => `${i} из ${total}${total > shown ? ` (показаны первые ${shown})` : ''}`,
  findCount: (total, shown) => `найдено: ${total}${total > shown ? ` (показаны первые ${shown})` : ''}`,
  findFailed: 'Поиск не удался',
  findOptions: 'Параметры',
  findMatchCase: 'Учитывать регистр',
  findWholeCell: 'Ячейка целиком',
  findLookIn: 'Область поиска',
  findInFormulas: 'формулы',
  findInValues: 'значения',
  findOrder: 'Просматривать',
  findByRows: 'по строкам',
  findByColumns: 'по колонкам',
  findWithin: (range) => `В выделении ${range}`,
  findWholeSheet: 'Искать по всему листу',
  replaceToggle: 'Заменить (Ctrl+H)',
  replacePlaceholder: 'Заменить на…',
  replaceOne: 'Заменить',
  replaceAll: 'Заменить все',
  replaceDone: (n) => (n ? `Сделано замен: ${n}` : 'Нечего заменить'),
  filter: 'Фильтр и сортировка',
  filterOn: 'Фильтр по этой колонке включён',
  filterColumn: (letters) => `Колонка ${letters}`,
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
    [`показано ${shown} из ${total} строк`, cols.length ? `фильтр по ${cols.join(', ')}` : '', sort ? `сортировка по ${sort.col} ${sort.dir === 'asc' ? '↑' : '↓'}` : '']
      .filter(Boolean)
      .join(', '),
  filterFailed: 'Не удалось применить фильтр',
  valuesFailed: 'Не удалось получить значения колонки',
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
  formatGeneral: 'Общий',
  formatNumber: 'Числовой',
  formatCurrency: 'Денежный',
  formatPercent: 'Процентный',
  formatDate: 'Краткая дата',
  formatTime: 'Время',
  formatText: 'Текстовый',
  formatOther: 'Другой',
}
