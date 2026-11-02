// The workbook's locale as a ref for the grid and the Number Format box: loaded into xlcalc once (`./locales`), with a late one discarded.
import { shallowRef, watch, type ShallowRef } from 'vue'
import type { Locale } from 'xlcalc/contract'
import { loadBookLocale } from './locales'

/**
 * The locale object of `tag()`, its tag registered in xlcalc from then on: null while its module loads; undefined for en-US,
 * a tag the package has no module for and a module that failed to load, which are treated as en-US. A module that arrives
 * after the tag changed is discarded.
 */
export function useLocaleModule(tag: () => string | undefined): ShallowRef<Locale | undefined | null> {
  const locale = shallowRef<Locale | undefined | null>(undefined)
  watch(
    tag,
    (t) => {
      const load = t && t !== 'en-US' ? loadBookLocale(t) : undefined
      locale.value = load ? null : undefined
      void load?.then((l) => {
        if (tag() === t) locale.value = l
      })
    },
    { immediate: true },
  )
  return locale
}
