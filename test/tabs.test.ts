// The sheet tabs as in Excel: shown with a single sheet too (the older behavior via `hideLoneTab`), hidden sheets omitted,
// and a "+" that inserts a sheet at the end with Excel's next default name and displays it.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, ruMessages } from '../src'
import type { EditOutcome, SheetInfo, SheetOp } from '../src/contract/source'
import { fakeSource, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const tabs = (w: VueWrapper) => w.findAll('.xs-tab').map((t) => t.text())
const one = () => memorySource({ sheets: [{ name: 'Sheet1', cells: { A1: 'a' } }] })

/** A source whose sheets grow with each insert; `refuse` rejects each insert. */
function book(names: string[], refuse = false) {
  let sheets: SheetInfo[] = names.map((name) => ({ name, maxRow: 2, maxCol: 2 }))
  const f = fakeSource(sheets)
  f.source.sheets = async () => sheets
  f.source.window = async (req) => ({ row: req.row, col: req.col, text: [[req.sheet]], values: [[req.sheet]], formulas: {}, colors: {}, styleIds: {}, styles: [], widths: [null], rowHeights: {}, merges: [] })
  const structure = vi.fn(async (op: SheetOp): Promise<EditOutcome> => {
    if (refuse) return { ok: false, reason: 'refused', message: 'Нельзя' }
    if (op.op === 'insertSheet') sheets = [...sheets, { name: op.name, maxRow: 0, maxCol: 0 }]
    if (op.op === 'deleteSheet') sheets = sheets.filter((x) => x.name !== op.sheet)
    return { ok: true, changes: [] }
  })
  f.source.structure = structure
  return { source: f.source, structure }
}

describe('tabs', () => {
  it('Control+Page Down and Page Up show the next and the previous shown sheet, as Excel, and stop at the ends', async () => {
    const src = memorySource({ sheets: [{ name: 'A', cells: { A1: 1 } }, { name: 'B', cells: {}, hidden: true }, { name: 'C', cells: {} }] })
    const w = await mountSheet(src)
    const on = () => w.find('.xs-tab-on').text()
    const press = async (k: string) => {
      await w.find('.xs-scroller').trigger('keydown', { key: k, ctrlKey: true })
      await settle()
    }
    await press('PageDown')
    expect(on()).toBe('C')
    await press('PageDown')
    expect(on()).toBe('C')
    await press('PageUp')
    expect(on()).toBe('A')
    await press('PageUp')
    expect(on()).toBe('A')
  })

  it('show with one sheet, as Excel; hideLoneTab hides them then', async () => {
    expect(tabs(await mountSheet(one()))).toEqual(['Sheet1'])
    expect((await mountSheet(one(), { hideLoneTab: true })).find('.xs-tabs').exists()).toBe(false)
    const two = memorySource({ sheets: [{ name: 'A', cells: {} }, { name: 'B', cells: {} }] })
    expect(tabs(await mountSheet(two, { hideLoneTab: true }))).toEqual(['A', 'B'])
  })

  it('"+" inserts a sheet at the end under Excel’s next name and shows it; a name taken goes on to the next', async () => {
    const b = book(['Sheet1', 'Sheet3'])
    const w = await mountSheet(b.source)
    await w.find('[data-xs-add-sheet]').trigger('click')
    await settle()
    expect(b.structure).toHaveBeenLastCalledWith({ op: 'insertSheet', name: 'Sheet4' })
    expect(w.find('.xs-tab-on').text()).toBe('Sheet4')
    await w.find('[data-xs-add-sheet]').trigger('click')
    await settle()
    expect(tabs(w)).toEqual(['Sheet1', 'Sheet3', 'Sheet4', 'Sheet5'])
  })

  it('names in the interface’s language, counting on past a sheet deleted meanwhile, as Excel does', async () => {
    const b = book(['Лист1'])
    const w = await mountSheet(b.source, { messages: ruMessages })
    await w.find('[data-xs-add-sheet]').trigger('click')
    await settle()
    await b.source.structure!({ op: 'deleteSheet', sheet: 'Лист2' })
    await w.find('[data-xs-add-sheet]').trigger('click')
    await settle()
    expect(b.structure.mock.calls.map((c) => c[0])).toEqual([
      { op: 'insertSheet', name: 'Лист2' },
      { op: 'deleteSheet', sheet: 'Лист2' },
      { op: 'insertSheet', name: 'Лист3' },
    ])
    expect(w.find('[data-xs-add-sheet]').attributes('title')).toBe(ruMessages.insertSheet)
  })

  it('has no "+" without structure or on a read-only sheet, and shows a refusal', async () => {
    expect((await mountSheet(one())).find('[data-xs-add-sheet]').exists()).toBe(false)
    expect((await mountSheet(book(['A']).source, { readOnly: true })).find('[data-xs-add-sheet]').exists()).toBe(false)
    const w = await mountSheet(book(['A'], true).source)
    await w.find('[data-xs-add-sheet]').trigger('click')
    await settle()
    expect(w.find('.xs-message').text()).toBe('Нельзя')
  })
})
