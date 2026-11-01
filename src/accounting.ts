// Excel for Windows' Accounting Number Format menu: its currencies and the codes each applies, as measured in Excel 16.0
// (build 20430) in en-US and ru-RU regional settings, each item applied to a cell and its NumberFormat read (01.11.2026).
import type { SheetMessages } from './contract/messages'

/** A currency of the menu, as the text `accountingCurrency` names it. */
export type AccountingCurrency = Parameters<SheetMessages['accountingCurrency']>[0]

type Item = { currency: AccountingCurrency; code: string }
const GBP: Item = { currency: 'gbp', code: '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-' }
const CNY: Item = { currency: 'cny', code: '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ;_ [$¥-804]* "-"??_ ;_ @_ ' }
const CHF: Item = { currency: 'chf', code: '_-* #,##0.00 [$CHF-100C]_-;-* #,##0.00 [$CHF-100C]_-;_-* "-"?? [$CHF-100C]_-;_-@_-' }
/**
 * Excel's list for an Excel whose interface and regional format are the workbook's locale, in its order: the locale's own
 * currency, then £, €, ¥ and CHF (measured in ru-RU; en-US's without the ruble a Russian Windows adds). The codes are those
 * Excel writes in that regional format: the euro follows the locale's negatives (en-US's parentheses); the pound's, the yuan's
 * and the franc's are the same in both.
 */
const LISTS: Readonly<Record<string, readonly Item[]>> = {
  'en-US': [
    { currency: 'usd', code: '_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "-"??_);_(@_)' },
    GBP,
    { currency: 'eur', code: '_([$€-2] * #,##0.00_);_([$€-2] * (#,##0.00);_([$€-2] * "-"??_);_(@_)' },
    CNY,
    CHF,
  ],
  'ru-RU': [
    { currency: 'rub', code: '_-* #,##0.00 [$₽-419]_-;-* #,##0.00 [$₽-419]_-;_-* "-"?? [$₽-419]_-;_-@_-' },
    GBP,
    { currency: 'eur', code: '_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "-"??_-;_-@_-' },
    CNY,
    CHF,
  ],
}

/** The menu's currencies for a workbook's locale, in order, each with its code; a locale the package has no module for is en-US's. */
export function accountingCurrencies(locale: string | undefined): readonly Item[] {
  return LISTS[locale ?? ''] ?? LISTS['en-US']!
}
