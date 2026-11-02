// Excel's Number Format box in the workbook's locale, as measured by xlcalc (`xlcalc/text`, without the engine): the items of
// its list and the category of a format, for the toolbar's list and the Format Cells dialog's Number tab; and the currencies
// of the Accounting Number Format button's menu.
import { computed, type ComputedRef } from 'vue'
import type { AccountingItem, FormatChoice, NumberFormatCategory } from 'xlcalc/contract'
import { accountingFormats, numberFormatCategory, numberFormatChoices } from 'xlcalc/text'
import { useLocaleModule } from './locale-ref'

export interface NumberFormats {
  /** The list's items in Excel's order, each with its name, the format it applies and that format's code; empty while the locale loads. */
  choices: ComputedRef<FormatChoice[]>
  /** The category the box shows for a code (General without one); null while the locale loads. */
  categoryOf: (code: string | undefined) => NumberFormatCategory | null
  /** Accounting Number Format's currencies in Excel's order, each with its name in the locale's language and its code; empty while the locale loads. */
  currencies: ComputedRef<AccountingItem[]>
}

/**
 * The box in the workbook's locale (`SheetView`'s `locale`, en-US without one): one of the package's locales loads as a
 * separate chunk, as the grid's formula text does; a tag the package has no module for is treated as en-US.
 */
export function useNumberFormats(tag: () => string | undefined): NumberFormats {
  /** Null while the locale's module loads; then its tag, loaded into xlcalc (undefined for en-US). */
  const module = useLocaleModule(tag)
  const loaded = computed(() => (module.value === null ? null : module.value?.tag))
  const choices = computed(() => (loaded.value === null ? [] : numberFormatChoices(loaded.value)))
  const currencies = computed(() => (loaded.value === null ? [] : accountingFormats(loaded.value).items))
  function categoryOf(code: string | undefined): NumberFormatCategory | null {
    if (loaded.value === null) return null
    try {
      return numberFormatCategory(code ?? 0, loaded.value)
    } catch {
      // a code still being typed that is not a valid format yet has no category
      return null
    }
  }
  return { choices, categoryOf, currencies }
}
