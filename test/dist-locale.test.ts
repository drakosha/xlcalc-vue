// The built main entry point in a ru-RU workbook: its bundled copy of xlcalc has a registry of locales of its own, which
// nothing but the component fills (a host's ./local registers ru-RU in the host's copy), so this runs dist/index.js and
// not src, where tests share one copy with ./local. Needs `npm run build` first, as dist.test.ts does.
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { LOCALES } from 'xlcalc'
import { settle, until } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

describe('dist/index.js in a ru-RU workbook', () => {
  it("takes the source's locale and steps the decimals, before and after its module arrives", async () => {
    // the package's own xlcalc, not the bundled one, which this test cannot see: it stays without ru-RU throughout
    expect(LOCALES.has('ru-RU')).toBe(false)
    const { SheetToolbar } = (await import('../dist/index.js')) as typeof import('../src')
    const format = vi.fn(async () => ({ ok: true as const, changes: [] }))
    const source = { locale: 'ru-RU', sheets: async () => [{ name: 'S', maxRow: 9, maxCol: 9 }], window: async () => ({ pending: { elapsedMs: 0 } }), format }
    const w = mount(SheetToolbar, { props: { source, state: { sheet: 'S', ref: 'A1', cellValue: 1.5 } }, attachTo: document.body })
    await w.find('[data-tb="decimal-up"]').trigger('click')
    await until(() => w.findAll('[data-tb="number"] option').some((o) => o.text() === 'Общий'), 5000)
    await w.find('[data-tb="decimal-up"]').trigger('click')
    await w.setProps({ state: { sheet: 'S', ref: 'A1', numberFormat: '0.00', cellValue: 1.5 } })
    await w.find('[data-tb="decimal-down"]').trigger('click')
    await settle()
    expect(format.mock.calls.map((c) => (c as unknown[])[1])).toEqual([
      [{ range: 'A1', numberFormat: '0.00' }],
      [{ range: 'A1', numberFormat: '0.00' }],
      [{ range: 'A1', numberFormat: '0.0' }],
    ])
    expect(LOCALES.has('ru-RU')).toBe(false)
  })
})
