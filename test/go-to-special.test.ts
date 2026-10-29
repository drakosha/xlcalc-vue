// Excel's Go To Special over `./local` (xlcalc 0.66's getSpecialCells) and the selection of several areas it makes: the
// dialog's options, the active cell kept or moved as Excel for Mac moved it, "No cells were found.", the rules' All and Same
// as Excel for Mac selected them, and what the sheet's commands do over several areas.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { SheetToolbar, type SheetViewApi, type SheetViewState } from '../src'
import type { SheetRules } from '../src/contract/workbook'
import { ruleCells } from '../src/core/special'
import { localSource, type LocalSource } from '../src/local'
import { address, active, cellText, clipboard, goTo, key, mountSheet, settle } from './helpers'

const open: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  for (const s of open.splice(0)) await s.close()
})

const v = (value: unknown) => ({ value })
/** The probe book measured in Excel for Mac: values B2 and L1:U3, a styled H9, row 12 of its own height, column J of its own width; two rules of each kind. */
const gt = (ranges: string[]) => ({ type: 'cellIs', operator: 'greaterThan', formulas: ['5'], ranges, style: { fill: { color: '#FFC7CE' } } })
const PROBE = {
  locale: 'en-US',
  styles: [{}, { font: { bold: true } }],
  sheets: [
    {
      name: 'P',
      cells: { B2: v(1), H9: { style: 1 }, L1: v('Region'), M1: v('Q1'), N1: v('Q2'), L2: v('north'), M2: v(1), N2: v(2), L3: v('south'), M3: v(3), N3: v(4), P1: v('Qty'), P2: v(1), P3: v(2), R1: v(2020), S1: v(2021), R2: v(5), S2: v(6), U1: v('a'), U2: v('b'), U3: v('c') },
      rows: { 12: { height: 30 } },
      columns: { J: { width: 20 } },
      conditionalFormats: [gt(['A1:A5']), gt(['A3:C3']), gt(['E1:E2'])],
      validations: [
        { range: 'D1:D3', type: 'whole', operator: 'between', formula1: '1', formula2: '10' },
        { range: 'D5', type: 'list', formula1: '"x,y"' },
        { range: 'F1:F2', type: 'whole', operator: 'between', formula1: '1', formula2: '10' },
      ],
    },
  ],
}
function probe(): LocalSource {
  const s = localSource(PROBE as never, { calc: 'inline' })
  open.push(s)
  return s
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
async function mounted(s: LocalSource): Promise<VueWrapper> {
  const w = await mountSheet(s)
  await settle(12)
  return w
}
/** Opens Go To Special, picks an option (and Same), presses OK. */
async function special(w: VueWrapper, kind: string, opts: { same?: boolean; values?: string[] } = {}): Promise<void> {
  expect(api(w).goToSpecial()).toBe(true)
  await settle()
  await w.find(`[data-xs-special-kind="${kind}"]`).setValue(true)
  if (opts.same) await w.find('[data-xs-special-same="same"]').setValue(true)
  if (opts.values) for (const box of w.findAll('[data-xs-special-value]')) await box.setValue(opts.values.includes(box.attributes('data-xs-special-value')!))
  await w.find('[data-xs-special] [data-xs-dialog-ok]').trigger('click')
  await settle(12)
}
/** The drawn cells of the selection's other areas and of its active range. */
const areaCells = (w: VueWrapper) => w.findAll('.xs-cell.xs-area').map((c) => c.attributes('data-cell'))
const selCells = (w: VueWrapper) => w.findAll('.xs-cell.xs-sel').map((c) => c.attributes('data-cell'))

describe("./local's specialCells", () => {
  it("finds as xlcalc does over a range, and Excel's last cell of the used range whatever the range (U12, as Excel for Mac answered)", async () => {
    const s = probe()
    expect(await s.specialCells('P', 'K1:N4', { type: 'blanks' })).toEqual(['K1:K3', 'K4:N4'])
    expect(await s.specialCells('P', 'L1:U3', { type: 'constants', values: ['text'] })).toEqual(['L1:N1', 'P1', 'U1:U3', 'L2:L3'])
    expect(await s.specialCells('P', 'A1', { type: 'lastCell' })).toEqual(['U12'])
    expect(await s.specialCells('P', 'C3:D4', { type: 'lastCell' })).toEqual(['U12'])
  })
})

describe("the rules' All and Same, as Excel for Mac selected them", () => {
  const rules = PROBE.sheets[0]! as unknown as SheetRules & { conditionalFormats: SheetRules['conditionalFormats'] }
  const at = (row: number, col: number) => ({ row, col })
  const addrs = (rects: { top: number; left: number; bottom: number; right: number }[]) => rects.map((r) => `${r.top}:${r.left}-${r.bottom}:${r.right}`)
  const book = { conditionalFormats: rules.conditionalFormats, validations: (PROBE.sheets[0] as unknown as SheetRules).validations }

  it("conditional formats: Same is the first rule over the active cell, the areas aside; All is every rule's cells within the areas, the sheet's for one cell", () => {
    expect(addrs(ruleCells(book, 'conditionalFormats', true, at(3, 1), null))).toEqual(['1:1-5:1'])
    expect(addrs(ruleCells(book, 'conditionalFormats', true, at(1, 5), null))).toEqual(['1:5-2:5'])
    expect(ruleCells(book, 'conditionalFormats', true, at(1, 7), null)).toEqual([])
    expect(addrs(ruleCells(book, 'conditionalFormats', true, at(3, 1), [{ top: 1, left: 1, bottom: 4, right: 2 }]))).toEqual(['1:1-5:1'])
    expect(addrs(ruleCells(book, 'conditionalFormats', false, at(1, 1), null))).toEqual(['1:1-5:1', '3:1-3:3', '1:5-2:5'])
    expect(addrs(ruleCells(book, 'conditionalFormats', false, at(1, 1), [{ top: 2, left: 1, bottom: 4, right: 2 }]))).toEqual(['2:1-4:1', '3:1-3:2'])
  })

  it('data validation: Same takes every rule with the active cell\'s settings; All within the areas', () => {
    expect(addrs(ruleCells(book, 'validation', true, at(1, 4), null))).toEqual(['1:4-3:4', '1:6-2:6'])
    expect(addrs(ruleCells(book, 'validation', true, at(5, 4), null))).toEqual(['5:4-5:4'])
    expect(ruleCells(book, 'validation', true, at(1, 7), null)).toEqual([])
    expect(addrs(ruleCells(book, 'validation', false, at(1, 4), null))).toEqual(['1:4-3:4', '5:4-5:4', '1:6-2:6'])
    expect(addrs(ruleCells(book, 'validation', false, at(1, 4), [{ top: 2, left: 3, bottom: 9, right: 5 }]))).toEqual(['2:4-3:4', '5:4-5:4'])
  })
})

describe('Go To Special', () => {
  it('selects the blanks of the selection as several areas, the active cell kept where one holds it (K1), else the first one\'s top-left (L4)', async () => {
    const w = await mounted(probe())
    await goTo(w, 'K1:N4')
    await special(w, 'blanks')
    expect(w.find('[data-xs-special]').exists()).toBe(false)
    expect([active(w), address(w)]).toEqual(['1:11', 'K1'])
    expect(selCells(w)).toEqual(['1:11', '2:11', '3:11'])
    expect(areaCells(w)).toEqual(['4:11', '4:12', '4:13', '4:14'])
    await goTo(w, 'L1:N4')
    expect(areaCells(w)).toEqual([])
    await special(w, 'blanks')
    expect([active(w), address(w)]).toEqual(['4:12', 'L4:N4'])
  })

  it('says "No cells were found." and leaves the selection as it was', async () => {
    const w = await mounted(probe())
    await goTo(w, 'V5:W6')
    await special(w, 'constants')
    expect(w.find('.xs-va').text()).toContain('No cells were found.')
    await w.find("[data-xs-alert-button='ok']").trigger('click')
    await settle()
    expect(address(w)).toBe('V5:W6')
  })

  it("finds constants by the kinds of their values, the current region, the last cell and the rules' cells", async () => {
    const w = await mounted(probe())
    await goTo(w, 'L1:N3')
    await special(w, 'constants', { values: ['numbers'] })
    expect(address(w)).toBe('M2:N3')
    await goTo(w, 'M2')
    await special(w, 'currentRegion')
    expect([address(w), active(w)]).toEqual(['L1:N3', '2:13'])
    await special(w, 'lastCell')
    expect(address(w)).toBe('U12')
    await goTo(w, 'D2')
    await special(w, 'validation', { same: true })
    expect([active(w), selCells(w), areaCells(w)]).toEqual(['2:4', ['1:4', '2:4', '3:4'], ['1:6', '2:6']])
  })

  it("opens from the Go To dialog's Special…, which closes Go To; not without the source's specialCells", async () => {
    const w = await mounted(probe())
    await key(w, 'g', { ctrlKey: true })
    await settle()
    await w.find('[data-xs-goto-special]').trigger('click')
    await settle()
    expect([w.find('[data-xs-goto]').exists(), w.find('[data-xs-special]').exists()]).toEqual([false, true])
    expect(w.findAll('[data-xs-special-kind]').map((x) => x.attributes('data-xs-special-kind'))).toEqual(['constants', 'formulas', 'blanks', 'currentRegion', 'lastCell', 'visible', 'conditionalFormats', 'validation'])
    const s = probe()
    const bare = await mounted({ ...s, specialCells: undefined } as unknown as LocalSource)
    expect(api(bare).goToSpecial()).toBe(false)
  })
})

describe('Go To Special and a sheet switch while it finds', () => {
  /** Two sheets over `./local`, whose specialCells answers on `release()`, rejecting an aborted call when it `honours` the signal. */
  function slow(honours: boolean) {
    const local = localSource({ sheets: [{ name: 'A', cells: { A1: v(1), C3: v(2), E5: v(3) } }, { name: 'B', cells: { A1: v('z') } }] } as never, { calc: 'inline' })
    open.push(local)
    let release = () => {}
    const source = Object.create(local) as LocalSource
    source.specialCells = (sheet, range, kind, signal) =>
      new Promise((resolve, reject) => {
        release = () => (honours && signal?.aborted ? reject(new DOMException('The operation was aborted.', 'AbortError')) : resolve(local.specialCells(sheet, range, kind)))
      })
    return { source, release: () => release() }
  }
  for (const honours of [false, true])
    it(`selects nothing on the sheet switched to and shows no message (a source that ${honours ? 'honours' : 'ignores'} the signal)`, async () => {
      const { source, release } = slow(honours)
      const w = await mounted(source)
      await goTo(w, 'A1:E5')
      expect(api(w).goToSpecial()).toBe(true)
      await settle()
      await w.find('[data-xs-special-kind="constants"]').setValue(true)
      await w.find('[data-xs-special] [data-xs-dialog-ok]').trigger('click')
      await settle()
      await w.findAll('.xs-tab')[1]!.trigger('click')
      await settle()
      release()
      await settle(12)
      expect(w.find('.xs-tab-on').text()).toBe('B')
      expect([areaCells(w), address(w), w.find('.xs-message').exists()]).toEqual([[], 'A1', false])
    })
})

describe('a selection of several areas', () => {
  /** The probe with the blanks of K1:N4 selected: K1:K3 (active K1) and K4:N4. */
  async function blanks(): Promise<{ s: LocalSource; w: VueWrapper }> {
    const s = probe()
    const w = await mounted(s)
    await goTo(w, 'K1:N4')
    await special(w, 'blanks')
    return { s, w }
  }

  it('Enter moves down the active area and on to the next one, Shift+Enter back; a click ends the areas', async () => {
    const { w } = await blanks()
    const state = () => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState
    expect([state().ref, state().areas]).toEqual(['K1:K3', ['K1:K3', 'K4:N4']])
    for (let i = 0; i < 3; i++) await key(w, 'Enter')
    await settle()
    expect([active(w), areaCells(w)]).toEqual(['4:11', ['1:11', '2:11', '3:11']])
    expect([state().ref, state().areas]).toEqual(['K4:N4', ['K1:K3', 'K4:N4']])
    await key(w, 'Enter', { shiftKey: true })
    await settle()
    expect(active(w)).toBe('3:11')
    await goTo(w, 'A1')
    expect(areaCells(w)).toEqual([])
    expect(state().areas).toBeUndefined()
  })

  it('Delete clears and Bold formats every area, as one step each', async () => {
    const s = probe()
    const w = await mounted(s)
    await goTo(w, 'L2:N3')
    await special(w, 'constants', { values: ['numbers'] })
    expect(address(w)).toBe('M2:N3')
    await goTo(w, 'L1:N3')
    await special(w, 'constants', { values: ['text'] })
    await key(w, 'b', { ctrlKey: true })
    await settle(12)
    const win = await s.window({ sheet: 'P', row: 1, rows: 3, col: 12, cols: 3 })
    if ('pending' in win) throw new Error('pending')
    const bold = (cell: string) => win.styles[win.styleIds[cell] ?? -1]?.font?.bold === true
    expect(['L1', 'M1', 'N1', 'L2', 'L3', 'M2'].map(bold)).toEqual([true, true, true, true, true, false])
    await key(w, 'Delete')
    await settle(12)
    expect([cellText(w, 1, 13), cellText(w, 3, 12), cellText(w, 2, 13)]).toEqual(['', '', '1'])
  })

  it("SheetToolbar's Bold and Format Cells' Border tab, through the sheet, format every area", async () => {
    const s = probe()
    const w = await mounted(s)
    await goTo(w, 'L1:N3')
    await special(w, 'constants', { values: ['text'] })
    const state = (w.emitted('state') ?? []).at(-1)![0] as SheetViewState
    const bar = mount(SheetToolbar, { props: { source: s, state, sheet: api(w) }, attachTo: document.body })
    await bar.find('[data-tb="bold"]').trigger('click')
    await settle(12)
    bar.unmount()
    expect(api(w).formatCells()).toBe(true)
    await settle()
    w.find('[data-fmt-tab="border"]').element.dispatchEvent(new Event('click'))
    await settle()
    await w.find('[data-fmt-edge="bottom"]').trigger('click')
    await w.find('[data-fmt] form, form[data-fmt]').trigger('submit')
    await settle(12)
    const win = await s.window({ sheet: 'P', row: 1, rows: 3, col: 12, cols: 3 })
    if ('pending' in win) throw new Error('pending')
    const style = (cell: string) => win.styles[win.styleIds[cell] ?? -1]
    expect(['L1', 'M1', 'N1', 'L3', 'M2'].map((c) => style(c)?.font?.bold === true)).toEqual([true, true, true, true, false])
    expect(['N1', 'L3', 'M2'].map((c) => style(c)?.border?.bottom?.style)).toEqual(['thin', 'thin', undefined])
  })

  it("copy, cut and paste refuse with Excel's message", async () => {
    const { w } = await blanks()
    clipboard(w, 'copy')
    await settle()
    expect(w.find('.xs-message').text()).toBe('This action won’t work on multiple selections.')
    expect(w.find('[data-xs-ants]').exists()).toBe(false)
  })
})
