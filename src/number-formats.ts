// Excel's Number Format box in the workbook's locale, as measured by xlcalc (`xlcalc/text`, without the engine): the items of
// its list and the category of a format, for the toolbar's list and the Format Cells dialog's Number tab.
import { computed, type ComputedRef } from 'vue'
import type { FormatChoice, NumberFormatCategory } from 'xlcalc/contract'
import { numberFormatCategory, numberFormatChoices } from 'xlcalc/text'
import { useLocaleModule } from './locale-ref'

export interface NumberFormats {
  /** The list's items in Excel's order, each with its name, the format it applies and that format's code; empty while the locale loads. */
  choices: ComputedRef<FormatChoice[]>
  /** The category the box shows for a code (General without one); null while the locale loads. */
  categoryOf: (code: string | undefined) => NumberFormatCategory | null
}

/**
 * The box in the workbook's locale (`SheetView`'s `locale`, en-US without one): one of the package's locales loads as a
 * separate chunk, as the grid's formula text does; a tag the package has no module for is treated as en-US.
 */
export function useNumberFormats(tag: () => string | undefined): NumberFormats {
  /** The locale object, undefined for en-US, null while its module loads. */
  const locale = useLocaleModule(tag)
  const choices = computed(() => (locale.value === null ? [] : numberFormatChoices(locale.value)))
  function categoryOf(code: string | undefined): NumberFormatCategory | null {
    if (locale.value === null) return null
    try {
      return numberFormatCategory(code ?? 0, locale.value)
    } catch {
      // a code still being typed that is not a valid format yet has no category
      return null
    }
  }
  return { choices, categoryOf }
}
