// Excel's Workbook Statistics: the rows Excel for Mac shows, the window opened by `workbookStatistics()` over the source's
// answer, and `./local` counting through xlcalc's `getWorkbookStatistics`.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookStatistics } from 'xlcalc/contract'
import { memorySource, ruMessages, type SheetViewApi } from '../src'
import { statisticsRows } from '../src/core/statistics'
import { closeAll, source } from './local-fixture'
import { mountSheet, settle } from './helpers'

afterEach(async () => {
  document.body.innerHTML = ''
  await closeAll()
})

const counts = { pivotTables: null, charts: null, images: null, formControls: null, objects: null, comments: null }
const STATS: WorkbookStatistics = {
  sheet: { end: 'C3', cells: 3, tables: 0, formulas: 2, notes: 1, ...counts },
  workbook: { sheets: 3, cells: 7, tables: 0, formulas: 4, notes: 2, ...counts },
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi

describe("the window's rows", () => {
  it('shows the end, cells, tables and formulas always, other counts above 0, never the workbook’s notes or what the workbook does not hold', () => {
    expect(statisticsRows(STATS)).toEqual({
      sheet: [['statisticsEnd', 'C3'], ['statisticsCells', '3'], ['statisticsTables', '0'], ['statisticsFormulas', '2'], ['statisticsNotes', '1']],
      workbook: [['statisticsSheets', '3'], ['statisticsCells', '7'], ['statisticsTables', '0'], ['statisticsFormulas', '4']],
    })
    const empty = statisticsRows({ sheet: { ...STATS.sheet, end: null, cells: 0, formulas: 0, notes: 0 }, workbook: { ...STATS.workbook, sheets: 1 } })
    expect(empty.sheet.map(([l]) => l)).toEqual(['statisticsCells', 'statisticsTables', 'statisticsFormulas'])
    // a source that counts PivotTables shows them where there are some
    expect(statisticsRows({ ...STATS, sheet: { ...STATS.sheet, pivotTables: 2 } as never }).sheet).toContainEqual(['statisticsPivotTables', '2'])
  })
})

describe('Workbook Statistics on the sheet', () => {
  it('opens with the source’s answer for the displayed sheet, and OK closes it', async () => {
    const workbookStatistics = vi.fn(async () => STATS)
    const w = await mountSheet({ ...memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }), workbookStatistics }, { messages: ruMessages })
    expect(api(w).workbookStatistics()).toBe(true)
    await settle()
    expect(workbookStatistics).toHaveBeenCalledWith('S', expect.any(AbortSignal))
    const box = w.find('[data-xs-statistics]')
    expect(box.find('.xs-modal-title').text()).toBe('Статистика книги')
    expect(box.find('[data-xs-stats="sheet"]').text()).toContain('Конец листаC3')
    expect(box.find('[data-xs-stats="workbook"]').text()).toContain('Листы3')
    await box.find('[data-xs-dialog-ok]').trigger('click')
    expect(w.find('[data-xs-statistics]').exists()).toBe(false)
  })

  it('is not offered without the source’s workbookStatistics', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }))
    expect(api(w).workbookStatistics()).toBe(false)
  })
})

describe('./local', () => {
  it('counts through xlcalc: data, formulas, a spill and notes', async () => {
    const s = source({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B2: { formula: 'SEQUENCE(2)' }, D4: { note: { text: 'n' } } } }, { name: 'T', cells: { A1: { value: 'x' } } }] })
    const got = await s.workbookStatistics!('S')
    expect(got.sheet).toMatchObject({ end: 'D4', cells: 3, formulas: 2, notes: 1, tables: 0 })
    expect(got.workbook).toMatchObject({ sheets: 2, cells: 4, formulas: 2 })
  })
})
