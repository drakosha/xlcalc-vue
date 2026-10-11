// The book's locale loaded into xlcalc before the session opens: en-US is built in, others are modules loaded on demand.
import { LOCALES, registerLocale, type Locale } from 'xlcalc'

/** Locale modules of xlcalc by tag, each a dynamic import a bundler can see. */
const MODULES: Record<string, () => Promise<Locale>> = {
  'ru-RU': () => import('xlcalc/locales/ru-RU').then((m) => m.ruRU),
}

/** Load the locale of `tag` once; a tag without a module is left to the session, which refuses it naming the locales it has. */
export async function loadLocale(tag: string): Promise<void> {
  if (LOCALES.has(tag)) return
  const load = MODULES[tag]
  if (load) registerLocale(await load())
}
