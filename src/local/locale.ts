// The book's locale loaded into xlcalc before the session opens: en-US is built in, others are modules loaded on demand.
import { LOCALES, registerLocale } from 'xlcalc'
import type { Locale } from 'xlcalc/contract'
import { LOCALE_MODULES } from '../locales'

/** Load the locale of `tag` once; a tag without a module is left to the session, which refuses it naming the locales it has. */
export async function loadLocale(tag: string): Promise<void> {
  if (LOCALES.has(tag)) return
  const load = LOCALE_MODULES[tag]
  if (load) registerLocale(await load())
}

/** The words of the functions' tips by tag (`xlcalc/function-text/*`), apart from the computing locale: only the tip and Insert Function take them. */
const FUNCTION_TEXT: Readonly<Record<string, () => Promise<Partial<Locale>>>> = {
  'ru-RU': () => import('xlcalc/function-text/ru-RU').then((m) => m.ruRUFunctionText),
}
const tips = new Map<string, Promise<Locale | string>>()

/** The loaded locale of `tag` with its tips' words joined, loaded once; the tag alone where it has none (en-US) or they fail to load. */
export function tipLocale(tag: string): Promise<Locale | string> {
  let got = tips.get(tag)
  if (!got) {
    const load = FUNCTION_TEXT[tag]
    const base = LOCALES.get(tag)
    got = load && base ? load().then((text): Locale | string => ({ ...base, ...text }), () => tag) : Promise.resolve(tag)
    tips.set(tag, got)
  }
  return got
}
