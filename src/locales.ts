// xlcalc's locale modules by tag, one list for the grid (formula text typed in a locale) and the local source (its session).
import type { Locale } from 'xlcalc/contract'

/** Each a dynamic import a bundler sees: the main entry carries it as a chunk of its own, `./local` leaves it to the host's xlcalc. */
export const LOCALE_MODULES: Readonly<Record<string, () => Promise<Locale>>> = {
  'ru-RU': () => import('xlcalc/locales/ru-RU').then((m) => m.ruRU),
}
