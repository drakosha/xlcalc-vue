// The workbook's locale loaded once per tag, one place for every component, the grid and the local source: importing one
// of xlcalc's locale modules loads it into the copy of xlcalc it is bundled with (xlcalc 0.77), so its tag works in every
// call from then on, in the main entry point's bundled xlcalc/a1 and xlcalc/text.
import type { Locale } from 'xlcalc/contract'

/** Each a dynamic import a bundler can see: the main entry point ships it as a separate chunk, `./local` leaves it to the host's xlcalc. */
const MODULES: Readonly<Record<string, () => Promise<Locale>>> = {
  'ru-RU': () => import('xlcalc/locales/ru-RU').then((m) => m.ruRU),
}

/** The tags loaded so far; not reactive, so that `./local` needs no vue on a server (`useLocaleModule` is the reactive view). */
const loaded = new Set<string>()
const loading = new Map<string, Promise<Locale | undefined>>()

/**
 * Loads the locale of `tag` once for all callers and gives its object, which the separately bundled `xlcalc/entry` chunk
 * and the locale's own words (separators, function names) still need. Undefined for en-US, built in, and for a tag the
 * package has no module for, which are treated as en-US; a module that fails to load rejects, and the next call tries again.
 */
export function loadBookLocale(tag: string): Promise<Locale | undefined> {
  const load = MODULES[tag]
  if (!load) return Promise.resolve(undefined)
  let got = loading.get(tag)
  if (!got) {
    got = load().then(
      (l) => (loaded.add(tag), l),
      (e: unknown) => {
        loading.delete(tag)
        throw e
      },
    )
    loading.set(tag, got)
  }
  return got
}

/** Whether `tag` is one of the package's locales other than en-US, which is loaded before its tag works. */
export const hasModule = (tag: string | undefined): boolean => !!tag && tag in MODULES

/**
 * `tag` as xlcalc's calls take it: itself once its module has loaded (en-US always), undefined (en-US) until then and for a
 * tag the package has no module for, so a call made before its module arrives does not throw.
 */
export const usableTag = (tag: string | undefined): string | undefined => (tag === 'en-US' || (tag && loaded.has(tag)) ? tag : undefined)
