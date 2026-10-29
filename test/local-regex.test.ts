// @vitest-environment node
// localSource's `regex` option: a level of xlcalc/regex (0.66.0) for match mode 3 of XLOOKUP and XMATCH, loaded on demand
// and registered before the workbook is calculated. The level is xlcalc's for every calculation of the page from then on.
import { describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import { closeAll, source, win } from './local-fixture'

// `(?i)` past the start is Excel's PCRE2 syntax, which the built-in level reads as JavaScript does: #VALUE!
const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 'abc' }, B1: { formula: 'XMATCH("a(?i)BC",A1,3)' } } }] })
// a pattern that backtracks for days in JavaScript: only pcreEngine's match limit ends it (#VALUE!)
const heavy = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { formula: 'XMATCH("(a+)+$",REPT("a",40)&"b",3)' } } }] })

describe('localSource: regular expressions', () => {
  it('without a level reads the pattern as JavaScript does', async () => {
    expect((await win(source(book()), 'S', 1, 1, 2, 1)).values).toEqual([[{ error: 'VALUE' }]])
    await closeAll()
  })

  it("with pcreSyntax reads Excel's syntax, and with pcreEngine ends a heavy pattern at Excel's match limit", async () => {
    expect((await win(source(book(), { regex: 'pcreSyntax' }), 'S', 1, 1, 2, 1)).values).toEqual([[1]])
    expect((await win(source(heavy(), { regex: 'pcreEngine' }), 'S', 1, 1, 1, 1)).values).toEqual([[{ error: 'VALUE' }]])
    await closeAll()
  })
})
