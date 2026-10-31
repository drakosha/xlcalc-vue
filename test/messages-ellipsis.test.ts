// Excel for Windows ends a command that opens a dialog with three dots ("Format Cells..."), not Excel for Mac's "…".
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { enMessages, ruMessages } from '../src'

/** The sheet's own progress texts and placeholders, which are not Excel's labels, keep "…". */
const OWN = ['loading', 'explainReading', 'pending', 'waiting', 'findPlaceholder', 'replacePlaceholder', 'filterText', 'filterSearch', 'filterReading', 'cfLoading']

/** The keys whose texts contain "…": strings at any depth, and functions by their source. */
function withEllipsis(messages: object): string[] {
  const keys = new Set<string>()
  const walk = (key: string, v: unknown): void => {
    if (typeof v === 'string' || typeof v === 'function') {
      if (String(v).includes('…')) keys.add(key)
    } else if (v && typeof v === 'object') for (const x of Object.values(v)) walk(key, x)
  }
  for (const [key, v] of Object.entries(messages)) walk(key, v)
  return [...keys].sort()
}

describe('message texts', () => {
  it('have no "…" but in the sheet\'s own texts, in English and in Russian', () => {
    expect(withEllipsis(enMessages)).toEqual([...OWN].sort())
    expect(withEllipsis(ruMessages)).toEqual([...OWN].sort())
  })

  it('have no "…" in their sources, tables included, but in the sheet\'s own texts', () => {
    const dir = join(import.meta.dirname, '..', 'src')
    const lines = readdirSync(dir)
      .filter((f) => /^messages.*\.ts$/.test(f))
      .flatMap((f) => readFileSync(join(dir, f), 'utf8').split('\n').filter((l) => l.includes('…')).map((l) => `${f}: ${l.trim()}`))
    expect(lines.filter((l) => !OWN.some((key) => l.includes(`: ${key}: `)))).toEqual([])
  })
})
