// The book's locale module as a ref for the grid and the Number Format box: loaded as its own chunk, a late one dropped.
import { shallowRef, watch, type ShallowRef } from 'vue'
import type { Locale } from 'xlcalc/contract'
import { LOCALE_MODULES } from './locales'

/**
 * The locale object of `tag()`: null while its module loads; undefined for a tag the package has no module for (en-US)
 * and for a module that failed to load, which reads as en-US. A module that comes after the tag changed is dropped.
 */
export function useLocaleModule(tag: () => string | undefined): ShallowRef<Locale | undefined | null> {
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
  return locale
}
