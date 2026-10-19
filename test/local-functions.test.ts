// @vitest-environment node
// localSource's names for Formula AutoComplete: every function xlcalc calculates, as typed in the workbook's locale (English
// where the locale has no name of its own), in the locale's alphabetical order.
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

describe('localSource: functionSignature', () => {
  it('in en-US: SUM’s tip, the name in any case', async () => {
    const sig = await source({ ...book(), locale: 'en-US' }).functionSignature('sum')
    expect(sig).toEqual({ name: 'SUM', fn: 'SUM', args: [{ name: 'number1', optional: false }, { name: 'number2', optional: true }], repeat: 1, text: 'SUM(number1, [number2], ...)', category: 'math', categoryName: 'Math & Trig', description: 'The sum of numbers.' })
  })

  it('in ru-RU: the Russian tip, by the Russian name or the English one; null for no function', async () => {
    const src = source()
    expect((await src.functionSignature('СУММ'))?.text).toBe('СУММ(число1;[число2];...)')
    expect((await src.functionSignature('sum'))?.text).toBe('СУММ(число1;[число2];...)')
    expect(await src.functionSignature('НЕТТАКОЙ')).toBeNull()
  })

  it('an aborted request rejects', async () => {
    const ctrl = new AbortController()
    ctrl.abort()
    await expect(source().functionSignature('СУММ', ctrl.signal)).rejects.toThrow(/aborted/)
  })
})
