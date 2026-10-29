// Excel's Create from Selection (Formulas > Defined Names, Ctrl+Shift+F3) over `./local` (xlcalc 0.66's
// getNamesFromSelection): the boxes Excel for Mac checks for a selection, the names made in one undo step, Excel's message
// for labels alone and its question before a name's definition is replaced.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { SheetViewApi } from '../src'
import type { SheetName } from '../src/contract/workbook'
import { localSource, type LocalSource } from '../src/local'
import { goTo, key, mountSheet, settle } from './helpers'

const open: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  for (const s of open.splice(0)) await s.close()
})

const v = (value: unknown) => ({ value })
/** The cells of the probe book measured in Excel for Mac, and the extra cases of the second probe. */
const CELLS = {
  L1: v('Region'), M1: v('Q1'), N1: v('Q2'), L2: v('north'), M2: v(1), N2: v(2), L3: v('south'), M3: v(3), N3: v(4),
  P1: v('Qty'), P2: v(1), P3: v(2), R1: v(2020), S1: v(2021), R2: v(5), S2: v(6), U1: v('a'), U2: v('b'), U3: v('c'),
  A1: v('h'), B1: v(5), C1: v('t'), A2: v(1), B2: v(2), C2: v(3), E1: v('h'), F1: v('t'), G1: v(7), E2: v(1), F2: v(2), G2: v(3),
  I1: v('h'), I2: v(5), I3: v('t'), J1: v(1), J2: v(2), J3: v(3), Z1: { formula: '"lbl"' }, AA1: { formula: '"x"&"y"' }, Z2: v(1), AA2: v(2),
}
function source(names: SheetName[] = []): LocalSource {
  const s = localSource({ locale: 'en-US', sheets: [{ name: 'P', cells: CELLS }], names: names.map((n) => ({ name: n.name, ref: n.ref })) } as never, { calc: 'inline' })
  open.push(s)
  return s
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
async function mounted(s: LocalSource, props: Record<string, unknown> = {}): Promise<VueWrapper> {
  const w = await mountSheet(s, props)
  await settle(12)
  return w
}
const SIDES = ['top', 'left', 'bottom', 'right'] as const
const dialog = (w: VueWrapper) => w.find('[data-xs-create-names]')
/** The boxes the dialog opened with for a selection. */
async function guessed(w: VueWrapper, range: string): Promise<string[]> {
  await goTo(w, range)
  expect(api(w).createNames()).toBe(true)
  await settle(12)
  const on = SIDES.filter((side) => (w.find(`[data-xs-create-names-side="${side}"]`).element as HTMLInputElement).checked)
  await w.find('[data-xs-create-names] [data-xs-dialog-cancel]').trigger('click')
  await settle()
  return on
}
const ok = async (w: VueWrapper) => {
  await w.find('[data-xs-create-names] [data-xs-dialog-ok]').trigger('click')
  await settle(16)
}
const alertText = (w: VueWrapper) => w.find('.xs-va').text()
async function answer(w: VueWrapper, button: 'yes' | 'no' | 'cancel' | 'ok'): Promise<void> {
  await w.find(`[data-xs-alert-button='${button}']`).trigger('click')
  await settle(16)
}
const named = async (s: LocalSource) => Object.fromEntries((await s.names()).map((n) => [n.name, n.ref]))

describe("./local's namesFromSelection", () => {
  it("gives xlcalc's names for the labels' sides, references in the workbook's locale", async () => {
    const s = source()
    expect(await s.namesFromSelection('P', 'L1:N3', { top: true, left: true })).toEqual([
      { name: 'Q1_', ref: 'P!$M$2:$M$3' },
      { name: 'Q2_', ref: 'P!$N$2:$N$3' },
      { name: 'north', ref: 'P!$M$2:$N$2' },
      { name: 'south', ref: 'P!$M$3:$N$3' },
      { name: 'Region', ref: 'P!$M$2:$N$3' },
    ])
  })
})

describe('Create from Selection', () => {
  it('checks the boxes Excel for Mac checked: Top row by the cell right of the corner, Left column by the cell below it', async () => {
    const w = await mounted(source())
    const cases: [string, string[]][] = [
      ['L1:N3', ['top', 'left']],
      ['P1:P3', ['top']],
      ['R1:S2', []],
      ['U1:U3', ['top']],
      ['L2:N3', ['left']],
      ['L1:N1', ['left']],
      ['A1:C2', []],
      ['E1:G2', ['top']],
      ['I1:J3', []],
      ['Z1:AA2', ['top']],
      ['L1', []],
    ]
    const got: [string, string[]][] = []
    for (const [range] of cases) got.push([range, await guessed(w, range)])
    expect(got).toEqual(cases)
  })

  it('makes the names in one undo step, and OK is off with no box checked', async () => {
    const s = source()
    const w = await mounted(s)
    await goTo(w, 'R1:S2')
    api(w).createNames()
    await settle(12)
    expect(w.find('[data-xs-create-names] [data-xs-dialog-ok]').attributes('disabled')).toBeDefined()
    await w.find('[data-xs-create-names] [data-xs-dialog-cancel]').trigger('click')
    await goTo(w, 'L1:N3')
    await key(w, 'F3', { ctrlKey: true, shiftKey: true })
    await settle(12)
    expect(dialog(w).exists()).toBe(true)
    await ok(w)
    expect(dialog(w).exists()).toBe(false)
    expect(await named(s)).toEqual({ Q1_: 'P!$M$2:$M$3', Q2_: 'P!$N$2:$N$3', north: 'P!$M$2:$N$2', south: 'P!$M$3:$N$3', Region: 'P!$M$2:$N$3' })
    await key(w, 'z', { ctrlKey: true })
    await settle(16)
    expect(await s.names()).toEqual([])
  })

  it("says Excel's message for labels alone, and makes nothing", async () => {
    const s = source()
    const w = await mounted(s)
    await goTo(w, 'L1:N1')
    api(w).createNames()
    await settle(12)
    await w.find('[data-xs-create-names-side="left"]').setValue(false)
    await w.find('[data-xs-create-names-side="top"]').setValue(true)
    await ok(w)
    expect(dialog(w).exists()).toBe(false)
    expect(alertText(w)).toContain('This selection isn’t valid.')
    await answer(w, 'ok')
    expect(await s.names()).toEqual([])
  })

  it("asks before replacing a name's other definition: Yes replaces, No keeps it, Cancel stops; the same definition is not asked about", async () => {
    const s = source([
      { name: 'Q1_', ref: 'P!$A$1' },
      { name: 'Q2_', ref: 'P!$A$2' },
      { name: 'north', ref: 'P!$M$2:$N$2' },
      { name: 'south', ref: 'P!$A$3' },
    ])
    const w = await mounted(s)
    await goTo(w, 'L1:N3')
    api(w).createNames()
    await settle(12)
    await ok(w)
    expect(alertText(w)).toContain("Replace existing definition of 'Q1_'?")
    await answer(w, 'yes')
    expect(alertText(w)).toContain("Replace existing definition of 'Q2_'?")
    await answer(w, 'no')
    expect(alertText(w)).toContain("Replace existing definition of 'south'?")
    await answer(w, 'cancel')
    expect(w.find('.xs-va').exists()).toBe(false)
    expect(await named(s)).toEqual({ Q1_: 'P!$M$2:$M$3', Q2_: 'P!$A$2', north: 'P!$M$2:$N$2', south: 'P!$A$3' })
  })

  it('is not offered without the source\'s namesFromSelection or on a read-only sheet', async () => {
    const s = source()
    const bare = await mounted({ ...s, namesFromSelection: undefined } as unknown as LocalSource)
    expect(api(bare).createNames()).toBe(false)
    document.body.innerHTML = ''
    const ro = await mounted(source(), { readOnly: true })
    expect(api(ro).createNames()).toBe(false)
  })
})
