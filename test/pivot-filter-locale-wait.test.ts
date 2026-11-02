// A PivotTable's Date Filters dialog before the workbook's locale has come: it opens once ru-RU's module has come and reads
// the date typed in ru-RU, rather than open with en-US's reading in its place.
import { describe, expect, it, vi } from 'vitest'
import { cellText, settle, until } from './helpers'
import { ok, q, sales, set } from './pivot-sheet'

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

async function pick(sel: string): Promise<void> {
  q<HTMLButtonElement>(sel)!.click()
  await settle()
}

describe("a PivotTable's filter dialog before the workbook's locale has come", () => {
  it('Date Filters waits for ru-RU and reads its date', async () => {
    // an en-US session under a sheet in ru-RU: the session opens without the module the sheet waits for
    const { w, s } = await sales({ rows: [4], values: [{ field: 3 }] }, 'R!A4', undefined, { props: { locale: 'ru-RU' } })
    await until(() => !!q('[data-pivot-button="A3"]'))
    await pick('[data-pivot-button="A3"]')
    await pick('[data-pfm="date"]')
    await pick('[data-pfm-op="after"]')
    await settle(12)
    expect(q('[data-pf-first]')).toBeNull()
    gate.open()
    await until(() => !!q('[data-pf-first]'))
    await set('[data-pf-first]', '01.01.2025')
    await ok()
    await until(() => cellText(w, 6, 1) === 'Grand Total')
    expect((await s.pivotTable('R', 'PivotTable1'))?.table.filters).toEqual([{ field: 4, type: 'dateNewerThan', args: ['2025-01-01T00:00:00'] }])
  })
})
