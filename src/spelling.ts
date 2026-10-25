// Excel's Spelling over a cell's text, shared by both sources: its words, those Excel skips by default, and the first misspelt one.
import type { SpellChecker } from './contract/workbook'

/** A word: letters, marks and digits, with apostrophes inside (`don't`). */
const WORD = /[\p{L}\p{M}\p{N}]+(?:['’][\p{L}\p{M}]+)*/gu

/** A word Excel's Spelling skips by default: one in uppercase or with digits. */
const skipped = (word: string): boolean => /\p{N}/u.test(word) || (word === word.toLocaleUpperCase() && word !== word.toLocaleLowerCase())

/** The first misspelt word of `text` starting at or after position `from`, skipping the lowercased words in `ignore`; null when there is none. */
export function misspelt(text: string, from: number, checker: SpellChecker, ignore: ReadonlySet<string>): { word: string; at: number } | null {
  for (const m of text.matchAll(WORD)) {
    const word = m[0]
    if (m.index < from || skipped(word) || ignore.has(word.toLocaleLowerCase())) continue
    if (!checker.check(word)) return { word, at: m.index }
  }
  return null
}
