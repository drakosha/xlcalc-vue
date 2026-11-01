// Excel for Windows' Accounting Number Format menu: its currencies and the codes each applies, as measured in Excel 16.0
// (build 20430) in en-US and ru-RU regional settings, each item applied to a cell and its NumberFormat read (01.11.2026).
import type { SheetMessages } from './contract/messages'

/** A currency of the menu, as the text `accountingCurrency` names it. */
export type AccountingCurrency = Parameters<SheetMessages['accountingCurrency']>[0]

/** The menu's fixed set in Excel's order; the workbook's locale's own currency comes first. */
const CURRENCIES: readonly AccountingCurrency[] = ['usd', 'rub', 'gbp', 'eur', 'cny', 'chf']

/**
 * The codes Excel writes for each item, by the regional settings it runs in: the dollar and the euro follow their negatives'
 * pattern (en-US's parentheses), the others are the same in both. The interface language does not change them.
 */
const EN_US: Readonly<Record<AccountingCurrency, string>> = {
  usd: '_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "-"??_);_(@_)',
  rub: '_-* #,##0.00 [$₽-419]_-;-* #,##0.00 [$₽-419]_-;_-* "-"?? [$₽-419]_-;_-@_-',
  gbp: '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-',
  eur: '_([$€-2] * #,##0.00_);_([$€-2] * (#,##0.00);_([$€-2] * "-"??_);_(@_)',
  cny: '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ;_ [$¥-804]* "-"??_ ;_ @_ ',
  chf: '_-* #,##0.00 [$CHF-100C]_-;-* #,##0.00 [$CHF-100C]_-;_-* "-"?? [$CHF-100C]_-;_-@_-',
}
const RU_RU: Readonly<Record<AccountingCurrency, string>> = {
  ...EN_US,
  usd: '_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* "-"??_ ;_-@_ ',
  eur: '_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "-"??_-;_-@_-',
}
/** By the workbook's locale: each its currency first and its codes; a locale the package has no module for is en-US's. */
const LOCALES: Readonly<Record<string, { own: AccountingCurrency; codes: Readonly<Record<AccountingCurrency, string>> }>> = {
  'en-US': { own: 'usd', codes: EN_US },
  'ru-RU': { own: 'rub', codes: RU_RU },
}

/** The menu's currencies for a workbook's locale, in order: its own currency, then Excel's set without it, each with its code. */
export function accountingCurrencies(locale: string | undefined): { currency: AccountingCurrency; code: string }[] {
  const { own, codes } = LOCALES[locale ?? ''] ?? LOCALES['en-US']!
  return [own, ...CURRENCIES.filter((c) => c !== own)].map((currency) => ({ currency, code: codes[currency] }))
}
