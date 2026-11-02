// The dialogs over a PivotTable whose boxes are read in the workbook's locale wait for its module: Grouping opened before
// ru-RU's module has come shows its dates in ru-RU once it comes, not in en-US's that its OK would then read in ru-RU; Value
// Field Settings keeps the name rather than rename it in en-US's words. The tests run in order, the module held until the last.
import { describe, expect, it, vi } from 'vitest'
import { settle, until } from './helpers'
import { menuAt, choose, q, sales, set } from './pivot-sheet'

/** ru-RU's module, held until the test lets it come. */
const gate = vi.hoisted(() => {
  let open: () => void = () => {}
  const opened = new Promise<void>((r) => (open = r))
  return { opened, open: () => open() }
})
vi.mock('xlcalc/locales/ru-RU', async (real) => {
  await gate.opened
  return real()
})

describe("a PivotTable's dialogs before the workbook's locale has come", () => {
  it("Value Field Settings keeps the name while ru-RU has not come, rather than take en-US's", async () => {
    const { w } = await sales({ rows: [4], values: [{ field: 3 }] }, 'R!A4', undefined, { props: { locale: 'ru-RU' } })
    await menuAt(w, '4:2')
    await choose('pivotValueSettings')
    await until(() => !!q('[data-pivot-value]'))
    await set('[data-pv-function]', 'count')
    expect(q<HTMLInputElement>('[data-pv-name]')!.value).toBe('Sum of Amount')
  })

  it('Grouping waits for ru-RU and opens with its dates', async () => {
    // an en-US session under a sheet in ru-RU: the session opens without the module the sheet waits for
    const { w } = await sales({ rows: [4], values: [{ field: 3 }] }, 'R!A4', undefined, { props: { locale: 'ru-RU' } })
    await menuAt(w, '4:1')
    await choose('pivotGroup')
    await settle(12)
    expect(q('[data-pivot-group]')).toBeNull()
    gate.open()
    await until(() => !!q('[data-pivot-group]'))
    expect(q<HTMLInputElement>('[data-pg-start]')!.value).toBe('01.04.2024')
  })
})
