// The example app's own texts, Russian and English; the sheet and the toolbar take the package's `ruMessages` or `enMessages`.

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
  lost: string
  lostMore: (n: number) => string
  lostSheets: (names: string[]) => string
  dropHere: string
  openFailed: (message: string) => string
  saveFailed: (message: string) => string
  opened: (ms: number) => string
  edited: (ms: number) => string
}

const ms = (n: number) => Math.round(n)

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
    lost: 'Not carried over from the file:',
    lostMore: (n) => ` (and ${n} more)`,
    lostSheets: (names) => `, ${names.length === 1 ? 'sheet' : 'sheets'} ${names.join(', ')}`,
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
    lost: 'Из файла не перенесено:',
    lostMore: (n) => ` (и ещё ${n})`,
    lostSheets: (names) => `, ${names.length === 1 ? 'лист' : 'листы'} ${names.join(', ')}`,
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
