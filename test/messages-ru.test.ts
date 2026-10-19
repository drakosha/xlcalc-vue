// The Russian texts name columns as Russian Excel does: "столбец", not "колонка".
import { describe, expect, it } from 'vitest'
import { ruMessages } from '../src'

/** Every Russian text: strings as they are, functions called with sample arguments. */
function texts(): [string, string][] {
  return Object.entries(ruMessages).flatMap(([key, v]): [string, string][] => {
    if (typeof v === 'string') return [[key, v]]
    if (typeof v !== 'function') return []
    try {
      return [[key, String((v as (...a: unknown[]) => unknown)(2, 5, true))]]
    } catch {
      return []
    }
  })
}

describe('Russian texts', () => {
  it('use "столбец" for a column', () => {
    expect(texts().filter(([, t]) => /колон/i.test(t))).toEqual([])
    expect(ruMessages.menuColumnWidth).toBe('Ширина столбца…')
    expect(ruMessages.filterColumn('B')).toBe('Столбец B')
    expect(ruMessages.freeze(0, 2, true)).toContain('2 столбца')
    expect(ruMessages.sheetTitle('S', 1, 1)).toBe('S: 1 строка, 1 столбец')
  })
})
