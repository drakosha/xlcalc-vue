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
  selectColumn: string
  selectRow: string
  selectAll: string
  resize: string
  conflict: string
  locked: (reason: string) => string
  editFailed: string
  tooManyCells: string
  pending: string
  close: string
  editHint: string
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
  selectColumn: 'Click selects the column',
  selectRow: 'Click selects the row',
  selectAll: 'Click selects the whole sheet',
  resize: 'Drag to change the width; a double click restores the book’s width',
  conflict: 'The book was changed elsewhere: check and press Enter again',
  locked: (reason) => `This cell cannot be changed: ${reason}`,
  editFailed: 'The edit did not go through',
  tooManyCells: 'Too many cells for one edit',
  pending: 'Saving…',
  close: 'Close',
  editHint: 'Click to edit the cell',
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
  selectColumn: 'Щелчок выделяет колонку',
  selectRow: 'Щелчок выделяет строку',
  selectAll: 'Щелчок выделяет весь лист',
  resize: 'Потяните, чтобы изменить ширину. Двойной щелчок вернёт ширину из книги',
  conflict: 'Книгу изменили в другом месте: проверьте и нажмите Enter ещё раз',
  locked: (reason) => `Эту ячейку менять нельзя: ${reason}`,
  editFailed: 'Правка не прошла',
  tooManyCells: 'Слишком много ячеек для одной правки',
  pending: 'Сохраняю…',
  close: 'Закрыть',
  editHint: 'Нажмите, чтобы править ячейку',
}
