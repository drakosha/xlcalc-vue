// The book's locale loaded into xlcalc before the session opens: en-US is built in, others are modules loaded on demand.
import { LOCALES, registerLocale } from 'xlcalc'
import { LOCALE_MODULES } from '../locales'

/** Load the locale of `tag` once; a tag without a module is left to the session, which refuses it naming the locales it has. */
export async function loadLocale(tag: string): Promise<void> {
  if (LOCALES.has(tag)) return
  const load = LOCALE_MODULES[tag]
  if (load) registerLocale(await load())
}
