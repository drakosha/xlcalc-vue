// The workbook's locale as a ref for the grid and the Number Format box: loaded into xlcalc once (`./locales`), with a late one discarded.
import { shallowRef, watch, type ShallowRef } from 'vue'
import type { Locale } from 'xlcalc/contract'
import { hasModule, loadBookLocale } from './locales'

/**
 * The locale object of `tag()`, its tag loaded into xlcalc from then on: null while its module loads and after it failed to
 * load, so that nothing is read or written as en-US in its place; undefined for en-US and a tag the package has no module
 * for, which are treated as en-US. A module that arrives after the tag changed is discarded.
 */
export function useLocaleModule(tag: () => string | undefined): ShallowRef<Locale | undefined | null> {
  const locale = shallowRef<Locale | undefined | null>(undefined)
  watch(
    tag,
    (t) => {
      const load = t && hasModule(t) ? loadBookLocale(t) : undefined
      locale.value = load ? null : undefined
      // a failed module leaves null: the next component that asks for it tries again
      load?.then(
        (l) => {
          if (tag() === t) locale.value = l
        },
        () => {},
      )
    },
    { immediate: true },
  )
  return locale
}
