// The example app's own texts, Russian and English; the sheet and the toolbar take the package's `ruMessages` or `enMessages`.
import type { XlsxIssueCode } from 'xlcalc/xlsx'

export type SpaLang = 'en' | 'ru'
export type BookLocale = 'en-US' | 'ru-RU'
export const BOOK_LOCALES: BookLocale[] = ['en-US', 'ru-RU']

export interface SpaTexts {
  newBook: string
  open: string
  saveAs: string
  settings: string
  bookName: string
  /** Excel's name of a new book: Book1, Книга1. */
  defaultName: (n: number) => string
  newTitle: string
  bookLocale: string
  bookLocaleHint: string
  uiLang: string
  create: string
  cancel: string
  apply: string
  close: string
  localeName: Record<BookLocale, string>
  langName: Record<SpaLang, string>
  ready: string
  busy: string
  saving: string
  /** The status while a file is read, and while its book computes for the first time. */
  reading: string
  computing: string
  /** A number as the interface writes it: 3 406, 3,406. */
  count: (n: number) => string
  /** The notice of what a file lost: its title, each kind's name, and its details. */
  lostTitle: string
  lossKind: Record<XlsxIssueCode, string>
  lostDetails: string
  lostHide: string
  lostOther: (n: number) => string
  lostSheets: (names: string[]) => string
  lostCells: (count: number, examples: string[]) => string
  lostNames: (count: number, examples: string[]) => string
  dropHere: string
  openFailed: (message: string) => string
  saveFailed: (message: string) => string
  opened: (ms: number) => string
  edited: (ms: number) => string
}

const ms = (n: number) => Math.round(n)
/** The first eight of a list, and how many more. */
const firstOf = (xs: string[], more: (n: number) => string) => xs.slice(0, 8).join(', ') + (xs.length > 8 ? more(xs.length - 8) : '')
const enCount = new Intl.NumberFormat('en-US')
const ruCount = new Intl.NumberFormat('ru-RU')

export const TEXTS: Record<SpaLang, SpaTexts> = {
  en: {
    newBook: 'New',
    open: 'Open',
    saveAs: 'Save as',
    settings: 'Book settings',
    bookName: 'Book name',
    defaultName: (n) => `Book${n}`,
    newTitle: 'New book',
    bookLocale: 'Book language',
    bookLocaleHint: 'Formulas, numbers and dates are typed and shown in it. Changing it opens the book again, and undo starts over.',
    uiLang: 'Interface language',
    create: 'Create',
    cancel: 'Cancel',
    apply: 'Apply',
    close: 'Close',
    localeName: { 'en-US': 'English (en-US)', 'ru-RU': 'Russian (ru-RU)' },
    langName: { en: 'English', ru: 'Русский' },
    ready: 'Ready',
    busy: 'Calculating…',
    saving: 'Saving…',
    reading: 'Reading the file…',
    computing: 'Computing the book…',
    count: (n) => enCount.format(n),
    lostTitle: 'Not carried over',
    lossKind: {
      externalLink: 'links to other workbooks',
      userFunction: 'user-defined functions',
      formula: 'formulas',
      dataTable: 'data tables',
      errorValue: 'unknown errors',
      errorConstant: 'errors typed as values',
      part: 'damaged parts',
      table: 'tables',
      pivotTable: 'pivot tables',
      drawing: 'pictures and charts',
      sheetKind: 'chart and macro sheets',
      cellMetadata: 'data types in cells',
      connection: 'data connections',
      macros: 'macros',
      conditionalFormat: 'conditional formatting',
      dataValidation: 'data validation',
      hyperlink: 'hyperlinks',
      comment: 'notes',
      richText: 'formatted text',
      style: 'formatting',
      numberFormat: 'number formats',
      defaultFont: 'default font',
      defaultWidth: 'default column width',
      frozenPanes: 'frozen panes',
      autoFilter: 'filters',
      outline: 'outline styles',
      sparkline: 'sparklines',
      print: 'print settings',
      protection: 'protection',
      calculation: 'calculation settings',
      hidden: 'hidden sheets',
      locale: 'book language',
      name: 'names',
      meta: 'cell data',
      textLength: 'long text',
      noCachedValues: 'stored values',
    },
    lostDetails: 'Details',
    lostHide: 'Hide details',
    lostOther: (n) => `and ${enCount.format(n)} more`,
    lostSheets: (names) => `${names.length === 1 ? 'Sheet' : 'Sheets'}: ${firstOf(names, (n) => ` and ${n} more`)}`,
    lostCells: (count, examples) => `${count === 1 ? 'Cell' : 'Cells'}: ${examples.join(', ')}${count > examples.length ? `, ${enCount.format(count)} in all` : ''}`,
    lostNames: (count, examples) => `${count === 1 ? 'Name' : 'Names'}: ${examples.join(', ')}${count > examples.length ? `, ${enCount.format(count)} in all` : ''}`,
    dropHere: 'Drop an xlsx file to open it',
    openFailed: (message) => `Could not open the file: ${message}`,
    saveFailed: (message) => `Could not save the book: ${message}`,
    opened: (n) => `opened in ${ms(n)} ms`,
    edited: (n) => `last change ${ms(n)} ms`,
  },
  ru: {
    newBook: 'Создать',
    open: 'Открыть',
    saveAs: 'Сохранить как',
    settings: 'Настройки книги',
    bookName: 'Имя книги',
    defaultName: (n) => `Книга${n}`,
    newTitle: 'Новая книга',
    bookLocale: 'Язык книги',
    bookLocaleHint: 'На нём вводятся и показываются формулы, числа и даты. Смена языка открывает книгу заново, история отмены начинается сначала.',
    uiLang: 'Язык интерфейса',
    create: 'Создать',
    cancel: 'Отмена',
    apply: 'Применить',
    close: 'Закрыть',
    localeName: { 'en-US': 'английский (en-US)', 'ru-RU': 'русский (ru-RU)' },
    langName: { en: 'English', ru: 'Русский' },
    ready: 'Готово',
    busy: 'Пересчёт…',
    saving: 'Сохранение…',
    reading: 'Чтение файла…',
    computing: 'Расчёт книги…',
    count: (n) => ruCount.format(n),
    lostTitle: 'Не перенесено',
    lossKind: {
      externalLink: 'связи с другими книгами',
      userFunction: 'пользовательские функции',
      formula: 'формулы',
      dataTable: 'таблицы данных',
      errorValue: 'неизвестные ошибки',
      errorConstant: 'ошибки, введённые значением',
      part: 'повреждённые части файла',
      table: 'таблицы',
      pivotTable: 'сводные таблицы',
      drawing: 'рисунки и диаграммы',
      sheetKind: 'листы диаграмм и макросов',
      cellMetadata: 'типы данных в ячейках',
      connection: 'подключения к данным',
      macros: 'макросы',
      conditionalFormat: 'условное форматирование',
      dataValidation: 'проверка данных',
      hyperlink: 'гиперссылки',
      comment: 'примечания',
      richText: 'форматирование текста',
      style: 'оформление',
      numberFormat: 'числовые форматы',
      defaultFont: 'шрифт по умолчанию',
      defaultWidth: 'ширина столбцов по умолчанию',
      frozenPanes: 'закрепление областей',
      autoFilter: 'фильтры',
      outline: 'стили структуры',
      sparkline: 'спарклайны',
      print: 'параметры печати',
      protection: 'защита',
      calculation: 'параметры вычислений',
      hidden: 'скрытые листы',
      locale: 'язык книги',
      name: 'имена',
      meta: 'данные ячеек',
      textLength: 'длинный текст',
      noCachedValues: 'сохранённые значения',
    },
    lostDetails: 'Подробнее',
    lostHide: 'Скрыть подробности',
    lostOther: (n) => `и ещё ${ruCount.format(n)}`,
    lostSheets: (names) => `${names.length === 1 ? 'Лист' : 'Листы'}: ${firstOf(names, (n) => ` и ещё ${n}`)}`,
    lostCells: (count, examples) => `${count === 1 ? 'Ячейка' : 'Ячейки'}: ${examples.join(', ')}${count > examples.length ? `, всего ${ruCount.format(count)}` : ''}`,
    lostNames: (count, examples) => `${count === 1 ? 'Имя' : 'Имена'}: ${examples.join(', ')}${count > examples.length ? `, всего ${ruCount.format(count)}` : ''}`,
    dropHere: 'Отпустите xlsx-файл, чтобы открыть его',
    openFailed: (message) => `Не удалось открыть файл: ${message}`,
    saveFailed: (message) => `Не удалось сохранить книгу: ${message}`,
    opened: (n) => `открыто за ${ms(n)} мс`,
    edited: (n) => `правка ${ms(n)} мс`,
  },
}

/** The interface language: `?lang=`, else the browser's, Russian for Russian and English for the rest. */
export function pickLang(asked: string | null): SpaLang {
  if (asked === 'ru' || asked === 'en') return asked
  return typeof navigator !== 'undefined' && /^ru\b/i.test(navigator.language) ? 'ru' : 'en'
}
