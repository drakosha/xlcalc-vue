// @vitest-environment node
// localSource's names for Formula AutoComplete: every function xlcalc computes, as the book's locale types it (English
// where the locale names none), in the locale's alphabetical order.
import { afterEach, describe, expect, it } from 'vitest'
import { FUNCTION_NAMES } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { book, closeAll, source } from './local-fixture'

afterEach(closeAll)

describe('localSource: functionNames', () => {
  it('in ru-RU: the Russian name, else the English one, each function once, by the Russian alphabet', async () => {
    const names = await source().functionNames()
    expect(names).toHaveLength(FUNCTION_NAMES.length)
    expect(names).toEqual(FUNCTION_NAMES.map((f) => ruRU.functions[f] ?? f).sort(new Intl.Collator('ru-RU').compare))
    expect(names).toContain('СУММ')
    expect(names).toContain('ABS')
    expect(names).not.toContain('SUM')
    expect(names.filter((n) => n.startsWith('СУ'))).toEqual(['СУММ', 'СУММЕСЛИ', 'СУММЕСЛИМН', 'СУММКВ', 'СУММКВРАЗН', 'СУММПРОИЗВ', 'СУММРАЗНКВ', 'СУММСУММКВ'])
    expect(names.filter((n) => n.startsWith('СТАНДОТКЛОН'))).toEqual(['СТАНДОТКЛОН', 'СТАНДОТКЛОН.В', 'СТАНДОТКЛОН.Г', 'СТАНДОТКЛОНА', 'СТАНДОТКЛОНП', 'СТАНДОТКЛОНПА'])
  })

  it('in en-US: the English names in order', async () => {
    const names = await source({ ...book(), locale: 'en-US' }).functionNames()
    expect(names).toContain('SUM')
    expect(names.filter((n) => n.startsWith('STDEV'))).toEqual(['STDEV', 'STDEV.P', 'STDEV.S', 'STDEVA', 'STDEVP', 'STDEVPA'])
    expect(names.indexOf('LOG')).toBeLessThan(names.indexOf('LOG10'))
  })

  it('an aborted request rejects', async () => {
    const ctrl = new AbortController()
    ctrl.abort()
    await expect(source().functionNames(ctrl.signal)).rejects.toThrow(/aborted/)
  })
})
