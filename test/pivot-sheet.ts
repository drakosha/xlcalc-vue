// SheetView over localSource with a PivotTable of salesBook on R, for the tests of the dialogs over a PivotTable: the
// sheet mounted, the cell menu opened, form controls set as a user sets them, and OK.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, expect } from 'vitest'
import type { PivotTablePatch } from 'xlcalc'
import type { SheetSource, SheetViewApi, SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { goTo, mountSheet, settle, until } from './helpers'
import { addSales, salesBook } from './pivot-fixture'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

/**
 * The sheet over R with PivotTable1 laid out at A3 by `definition`, the active cell at `at`: in `locale` (the sheet's too),
 * with the sheet's `props`, its source as `wrap` gives it; it waits for the PivotTable's commands unless `editable` is false.
 */
export async function sales(
  definition: PivotTablePatch,
  at = 'R!B4',
  locale?: string,
  opts: { props?: Record<string, unknown>; wrap?: (s: LocalSource) => SheetSource; editable?: boolean } = {},
): Promise<{ w: VueWrapper; s: LocalSource }> {
  const s = localSource(salesBook(locale), { calc: 'inline' })
  sources.push(s)
  expect(await s.structure(addSales(definition))).toMatchObject({ ok: true })
  const w = await mountSheet(opts.wrap ? opts.wrap(s) : s, { style: 'height: 500px', pivotFields: false, ...(locale ? { locale } : {}), ...opts.props })
  await goTo(w, at)
  if (opts.editable !== false) await until(() => lastState(w)?.canEditPivot === true)
  else await until(() => !!lastState(w)?.pivot)
  return { w, s }
}
export const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
export const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
export const q = <T extends HTMLElement = HTMLElement>(sel: string) => document.querySelector<T>(sel)
export async function menuAt(w: VueWrapper, cell: string): Promise<void> {
  await w.find(`[data-cell="${cell}"]`).trigger('contextmenu', { button: 2, clientX: 40, clientY: 60 })
  await settle()
}
export const menuItem = (id: string) => q<HTMLButtonElement>(`[data-xs-menu="${id}"]`)
export async function choose(id: string, sub?: string): Promise<void> {
  if (sub) {
    menuItem(sub)!.click()
    await settle()
  }
  menuItem(id)!.click()
  await settle()
}
/** Sets a form control's value as a user would. */
export async function set(sel: string, value: string | boolean | string[]): Promise<void> {
  const el = q<HTMLInputElement & HTMLSelectElement>(sel)!
  if (Array.isArray(value)) for (const o of el.options) o.selected = value.includes(o.value)
  else if (typeof value === 'boolean') el.checked = value
  else el.value = value
  const kind = (el as HTMLInputElement).type
  el.dispatchEvent(new Event(kind === 'checkbox' || kind === 'radio' || el.tagName === 'SELECT' ? 'change' : 'input'))
  await settle()
}
export async function ok(): Promise<void> {
  q<HTMLButtonElement>('[data-xs-dialog-ok]')!.click()
  await settle()
}
