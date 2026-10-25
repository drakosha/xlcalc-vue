// Excel's Symbol dialog data: the Subset list's Unicode blocks, the characters of the grid, the Special Characters tab and the recently used symbols of a new install.
import type { SymbolSpecial, SymbolSubset } from './contract/messages'

/** The Subset list: Unicode blocks a text font usually covers, in code order, by Excel's names (the messages' `symbolSubset`). */
export const SYMBOL_SUBSETS: readonly { key: SymbolSubset; from: number; to: number }[] = [
  { key: 'basicLatin', from: 0x20, to: 0x7f },
  { key: 'latin1', from: 0xa0, to: 0xff },
  { key: 'latinExtA', from: 0x100, to: 0x17f },
  { key: 'latinExtB', from: 0x180, to: 0x24f },
  { key: 'ipa', from: 0x250, to: 0x2af },
  { key: 'spacingModifiers', from: 0x2b0, to: 0x2ff },
  { key: 'combiningMarks', from: 0x300, to: 0x36f },
  { key: 'greek', from: 0x370, to: 0x3ff },
  { key: 'cyrillic', from: 0x400, to: 0x4ff },
  { key: 'cyrillicSupplement', from: 0x500, to: 0x52f },
  { key: 'phoneticExt', from: 0x1d00, to: 0x1d7f },
  { key: 'latinExtAdditional', from: 0x1e00, to: 0x1eff },
  { key: 'greekExt', from: 0x1f00, to: 0x1fff },
  { key: 'punctuation', from: 0x2000, to: 0x206f },
  { key: 'superSub', from: 0x2070, to: 0x209f },
  { key: 'currency', from: 0x20a0, to: 0x20cf },
  { key: 'letterlike', from: 0x2100, to: 0x214f },
  { key: 'numberForms', from: 0x2150, to: 0x218f },
  { key: 'arrows', from: 0x2190, to: 0x21ff },
  { key: 'math', from: 0x2200, to: 0x22ff },
  { key: 'technical', from: 0x2300, to: 0x23ff },
  { key: 'enclosed', from: 0x2460, to: 0x24ff },
  { key: 'boxDrawing', from: 0x2500, to: 0x257f },
  { key: 'blocks', from: 0x2580, to: 0x259f },
  { key: 'geometric', from: 0x25a0, to: 0x25ff },
  { key: 'miscSymbols', from: 0x2600, to: 0x26ff },
  { key: 'dingbats', from: 0x2700, to: 0x27bf },
  { key: 'latinExtC', from: 0x2c60, to: 0x2c7f },
  { key: 'presentationForms', from: 0xfb00, to: 0xfb4f },
  { key: 'specials', from: 0xfff0, to: 0xffff },
]

/** A code point the grid shows: assigned, and neither a control character nor a surrogate. */
const SHOWN = /^[^\p{Cn}\p{Cc}\p{Cs}]$/u

let codes: readonly number[] | null = null
/** The grid's characters: every shown code point of the subsets, in code order, as Excel's grid runs through the font; built on first use. */
export function symbolCodes(): readonly number[] {
  codes ??= SYMBOL_SUBSETS.flatMap(({ from, to }) => {
    const out: number[] = []
    for (let c = from; c <= to; c++) if (SHOWN.test(String.fromCodePoint(c))) out.push(c)
    return out
  })
  return codes
}

/** The subset a code point belongs to, or null outside them. */
export function subsetOf(code: number): SymbolSubset | null {
  return SYMBOL_SUBSETS.find((s) => code >= s.from && code <= s.to)?.key ?? null
}

/** A code point as Excel's Character code box shows it: hex, four digits at least (`00A9`). */
export const codeHex = (code: number): string => code.toString(16).toUpperCase().padStart(4, '0')

/** The code point a Character code box's text names (hex, an optional `U+`), when it is a character one can insert; otherwise null. */
export function parseCode(text: string): number | null {
  const m = /^\s*(?:U\+)?([0-9a-f]{1,6})\s*$/i.exec(text)
  if (!m) return null
  const code = parseInt(m[1]!, 16)
  return code <= 0x10ffff && SHOWN.test(String.fromCodePoint(code)) ? code : null
}

/** Excel's Special Characters tab, in its order. */
export const SYMBOL_SPECIALS: readonly { key: SymbolSpecial; code: number }[] = [
  { key: 'emDash', code: 0x2014 },
  { key: 'enDash', code: 0x2013 },
  { key: 'nonBreakingHyphen', code: 0x2011 },
  { key: 'optionalHyphen', code: 0xad },
  { key: 'emSpace', code: 0x2003 },
  { key: 'enSpace', code: 0x2002 },
  { key: 'quarterEmSpace', code: 0x2005 },
  { key: 'nonBreakingSpace', code: 0xa0 },
  { key: 'copyright', code: 0xa9 },
  { key: 'registered', code: 0xae },
  { key: 'trademark', code: 0x2122 },
  { key: 'section', code: 0xa7 },
  { key: 'paragraph', code: 0xb6 },
  { key: 'ellipsis', code: 0x2026 },
  { key: 'singleOpeningQuote', code: 0x2018 },
  { key: 'singleClosingQuote', code: 0x2019 },
  { key: 'doubleOpeningQuote', code: 0x201c },
  { key: 'doubleClosingQuote', code: 0x201d },
  { key: 'noWidthOptionalBreak', code: 0x200c },
  { key: 'noWidthNonBreak', code: 0x200d },
]

/** How many recently used symbols the dialog shows, and the row of a new install, as Office's. */
export const RECENT_SYMBOLS = 16
export const DEFAULT_RECENT_SYMBOLS: readonly number[] = [0x20ac, 0xa3, 0xa5, 0xa9, 0xae, 0x2122, 0xb1, 0x2260, 0x2264, 0x2265, 0xf7, 0xd7, 0x221e, 0xb5, 0x3b1, 0x3b2]
