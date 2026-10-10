// What the edit field opens with for a value: a number in the book's locale, a text with Excel's apostrophe where needed.

/** A number's edit text in the book's locale: the full value, a percent as a percent, a date or time as shown. */
export function numberEditText(value: number, shown: string, locale: string): string {
  const raw = (n: number) => {
    try {
      return new Intl.NumberFormat(locale, { useGrouping: false, maximumSignificantDigits: 15 }).format(n)
    } catch {
      return String(n)
    }
  }
  const t = shown.trim()
  if (/^-?[\d\s.,  ]+%$/.test(t)) return `${raw(value * 100)}%`
  if (/^\d{1,4}[./-]\d{1,2}([./-]\d{1,4})?(\s+\d{1,2}:\d{2}(:\d{2})?)?$/.test(t) || /^\d{1,2}:\d{2}(:\d{2})?(\s?[AP]M)?$/i.test(t)) return t
  return raw(value)
}

// What a book reads as a value when typed, in ru-RU or en-US alike: an apostrophe too many is harmless, a text turned
// into a number on Enter is lost, so either locale's forms count in both.
const DIGITS = String.raw`(?:\d{1,3}(?:[,.\s  ]\d{3})+|\d+)(?:[.,]\d+)?|[.,]\d+`
const MONEY = String.raw`[$€£¥₽]`
const NUMBER = new RegExp(String.raw`^[+-]?\(?\s*(?:${MONEY}\s*)?(?:${DIGITS})(?:[eE][+-]?\d+)?\s*(?:%|${MONEY}|р\.|руб\.?)?\s*\)?$`, 'u')
const TIME = String.raw`\d{1,2}:\d{2}(?::\d{2}(?:[.,]\d+)?)?(?:\s?[ap]\.?m\.?)?`
const DATE = new RegExp(String.raw`^\d{1,4}[./-]\d{1,2}(?:[./-]\d{1,4})?(?:\s+${TIME})?$|^${TIME}$`, 'iu')
const MONTH = String.raw`(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|янв|фев|мар|апр|ма[йя]|июн|июл|авг|сен|окт|ноя|дек)\p{L}*\.?`
const NAMED_DATE = new RegExp(String.raw`^\d{1,2}[\s-]+${MONTH}(?:[\s-]+\d{2,4})?$|^${MONTH}[\s-]+\d{1,4}(?:,?\s+\d{2,4})?$`, 'iu')
const LOGICAL = /^(?:true|false|истина|ложь)$/iu
const ERROR = /^#(?:N\/A|Н\/Д|[^\s#]+[!?])$/iu

/** Whether a book would read the text as a number, percent, date, time, boolean, error or formula when typed. */
export function readsAsValue(text: string): boolean {
  const t = text.trim()
  if (!t) return false
  if (/^[=+\-@]./u.test(t) || t.startsWith("'")) return true
  return NUMBER.test(t) || DATE.test(t) || NAMED_DATE.test(t) || LOGICAL.test(t) || ERROR.test(t)
}

/** A text's edit text: as it is, with Excel's apostrophe when the book would read it as something else. */
export function textEditText(text: string): string {
  return readsAsValue(text) ? `'${text}` : text
}
