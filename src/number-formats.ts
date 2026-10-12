// Excel's Number Format box in the book's locale, as xlcalc measured it (`xlcalc/text`, without the engine): the items of
// its list and the category a format shows by, for the toolbar's list and the Format Cells dialog's Number tab.
import { computed, shallowRef, watch, type ComputedRef } from 'vue'
import type { FormatChoice, Locale, NumberFormatCategory } from 'xlcalc/contract'
import { numberFormatCategory, numberFormatChoices } from 'xlcalc/text'
import { LOCALE_MODULES } from './locales'

export interface NumberFormats {
  /** The list's items in Excel's order, each with its name, the format it applies and that format's code; empty while the locale loads. */
  choices: ComputedRef<FormatChoice[]>
  /** The category the box shows for a code (General without one); null while the locale loads. */
  categoryOf: (code: string | undefined) => NumberFormatCategory | null
}

/**
 * The box in the book's locale (`SheetView`'s `locale`, en-US without one): a locale of the package's loads as its own
 * chunk, as the grid's formula text does; a tag the package has no module for is read as en-US.
 */
export function useNumberFormats(tag: () => string | undefined): NumberFormats {
  /** The locale object, undefined for en-US, null while its module loads. */
  const locale = shallowRef<Locale | undefined | null>(undefined)
  watch(
    tag,
    (t) => {
      const load = t ? LOCALE_MODULES[t] : undefined
      locale.value = load ? null : undefined
      if (!load) return
      const take = (l: Locale | undefined) => {
        if (tag() === t) locale.value = l
      }
      void load().then(take, () => take(undefined))
    },
    { immediate: true },
  )
  const choices = computed(() => (locale.value === null ? [] : numberFormatChoices(locale.value)))
  function categoryOf(code: string | undefined): NumberFormatCategory | null {
    if (locale.value === null) return null
    try {
      return numberFormatCategory(code ?? 0, locale.value)
    } catch {
      // a code being typed that is no format yet has no category
      return null
    }
  }
  return { choices, categoryOf }
}
