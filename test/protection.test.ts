// A protected sheet on SheetView: Excel's alert before an edit of a locked cell, hidden formulas, the commands Excel
// disables, Protect Sheet and Unprotect Sheet in the tab menu and as methods, and Tab among the cells that can be selected.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput, SheetProtection } from 'xlcalc'
import type { SheetViewApi, SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { active, barField, editorOpen, goTo, key, mountSheet, paste, settle, until } from './helpers'

const ALERT = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password."

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

/** A form: B1:B3 and D1 unlocked, A2's formula hidden; protected with `protection`, or not with null. */
function form(protection: SheetProtection | null = {}): LocalSource {
  const book: BookInput = {
    styles: [{}, { protection: { locked: false } }, { protection: { hidden: true } }],
    sheets: [{ name: 'Form', cells: { A1: { value: 'Name' }, B1: { value: 'Ann', style: 1 }, A2: { formula: 'LEN(B1)', style: 2 }, B2: { value: 'x', style: 1 }, B3: { value: 'y', style: 1 }, D1: { value: 'z', style: 1 } }, ...(protection ? { protection } : {}) }],
  }
  const s = localSource(book, { calc: 'inline' })
  sources.push(s)
  return s
}
async function mounted(source: LocalSource): Promise<VueWrapper> {
  const w = await mountSheet(source)
  await until(() => w.find('[data-cell="1:1"]').text() === 'Name')
  return w
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
const alertText = () => document.querySelector('[data-xs-alert="notice"] [data-xs-alert-message]')?.textContent ?? null
const menuItem = (id: string) => document.querySelector<HTMLButtonElement>(`[data-xs-menu="${id}"]`)
async function closeAlert(): Promise<void> {
  document.querySelector<HTMLButtonElement>('[data-xs-alert-button="ok"]')?.click()
  await settle()
}

describe('a protected sheet', () => {
  it("shows Excel's alert instead of opening the editor on a locked cell, and edits an unlocked one", async () => {
    const w = await mounted(form())
    await goTo(w, 'A1')
    await key(w, 'q')
    await settle()
    expect(alertText()).toBe(ALERT)
    expect(editorOpen(w)).toBe(false)
    await closeAlert()
    await goTo(w, 'B1')
    await key(w, 'q')
    await settle()
    expect(alertText()).toBeNull()
    expect(editorOpen(w)).toBe(true)
  })

  it("refuses Delete and a paste over a locked cell with the alert, before anything is sent", async () => {
    const source = form()
    const w = await mounted(source)
    await goTo(w, 'A1:B1')
    await key(w, 'Delete')
    await settle()
    expect(alertText()).toBe(ALERT)
    await closeAlert()
    await goTo(w, 'B1')
    paste(w, 'a\tb')
    await settle()
    expect(alertText()).toBe(ALERT)
    await closeAlert()
    expect((await source.book()).sheets[0]!.cells.B1?.value).toBe('Ann')
  })

  it("keeps a hidden formula out of the formula bar and the state; the cell shows its value", async () => {
    const source = form()
    const w = await mounted(source)
    await goTo(w, 'A2')
    expect(barField(w).value).toBe('')
    expect(lastState(w)?.formula).toBeUndefined()
    expect(w.find('[data-cell="2:1"]').text()).toBe('3')
    expect(lastState(w)?.protection).toEqual({})
    await api(w).unprotectSheet()
    await until(() => barField(w).value === '=LEN(B1)')
    expect(lastState(w)?.protection).toBeUndefined()
  })

  it('keeps a hidden formula out of the tooltip, Show Formulas (an empty cell, as in Excel for Mac) and formula auditing', async () => {
    const w = await mounted(form())
    expect(w.find('[data-cell="2:1"]').attributes('title')).toBe('3')
    api(w).showFormulas(true)
    await settle()
    expect(w.find('[data-cell="2:1"]').text()).toBe('')
    expect(w.find('[data-cell="1:2"]').text()).toBe('Ann')
    await goTo(w, 'A2')
    // Excel disables Trace Precedents and Trace Dependents on a protected sheet
    expect(api(w).explain()).toBe(false)
    await goTo(w, 'B1')
    expect(await api(w).traceDependents()).toBe(false)
    await w.find('[data-cell="2:1"]').trigger('contextmenu', { button: 2, clientX: 200, clientY: 100 })
    await settle()
    expect(menuItem('precedents')).toBeNull()
    document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    await api(w).unprotectSheet()
    await until(() => w.find('[data-cell="2:1"]').text() === '=LEN(B1)')
    expect(w.find('[data-cell="2:1"]').attributes('title')).toBe('=LEN(B1)\n3')
    await goTo(w, 'B1')
    expect(await api(w).traceDependents()).toBe(true)
    await goTo(w, 'A2')
    expect(api(w).explain()).toBe(true)
  })

  it("shows the source's refusal on a protected sheet as Excel's alert", async () => {
    const w = await mounted(form({ deleteRows: true }))
    const out = api(w).structure({ op: 'deleteRows', sheet: 'Form', row: 1 })
    await until(() => alertText() !== null)
    expect(alertText()).toMatch(/^You are trying to delete a row that contains a locked cell\./)
    await closeAlert()
    expect(await out).toMatchObject({ ok: false, code: 'protected' })
    expect(w.find('.xs-message').exists()).toBe(false)
  })

  it('disables in the row menu what the protection does not allow, and Group and Ungroup always', async () => {
    const w = await mounted(form({ insertRows: true }))
    await w.find('.xs-num').trigger('contextmenu', { button: 2, clientX: 20, clientY: 40 })
    await settle()
    expect(['insert', 'delete', 'size', 'hide', 'unhide', 'group', 'ungroup'].map((id) => menuItem(id)?.disabled)).toEqual([false, true, true, true, true, true, true])
  })

  it('opens Format Cells only with Format cells, and its Protection tab keeps Locked and Hidden as they are', async () => {
    const off = await mounted(form())
    expect(api(off).formatCells()).toBe(false)
    await goTo(off, 'B1:B2')
    expect(await api(off).merge('mergeCells')).toBeNull()
    off.unmount()
    const on = await mounted(form({ formatCells: true }))
    await goTo(on, 'B1')
    expect(api(on).formatCells()).toBe(true)
    await settle()
    const locked = document.querySelector<HTMLInputElement>('[data-fmt-locked]')!
    expect([locked.checked, locked.disabled, document.querySelector<HTMLInputElement>('[data-fmt-hidden]')!.disabled]).toEqual([false, true, true])
  })

  it('Tab moves among the unlocked cells when locked ones cannot be selected; a click on a locked one does nothing', async () => {
    const w = await mounted(form({ selectLockedCells: false }))
    await goTo(w, 'B1')
    await key(w, 'Tab')
    await settle()
    expect(active(w)).toBe('1:4')
    await key(w, 'Tab')
    await settle()
    expect(active(w)).toBe('2:2')
    await key(w, 'Tab')
    await key(w, 'Tab')
    await settle()
    expect(active(w)).toBe('1:2')
    // an arrow goes over locked cells to the next unlocked one that way, and stays put when there is none
    await key(w, 'ArrowRight')
    await settle()
    expect(active(w)).toBe('1:4')
    await key(w, 'ArrowDown')
    await settle()
    expect(active(w)).toBe('1:4')
    await w.find('[data-cell="1:1"]').trigger('pointerdown', { button: 0 })
    await settle()
    expect(active(w)).toBe('1:4')
  })
})

describe('Protect Sheet and Unprotect Sheet', () => {
  it("protects the sheet from its tab's menu with the boxes checked, then unprotects it", async () => {
    const source = form(null)
    const w = await mounted(source)
    await w.find('.xs-tab').trigger('contextmenu', { button: 2, clientX: 20, clientY: 400 })
    await settle()
    expect(menuItem('unprotectSheet')).toBeNull()
    menuItem('protectSheet')!.click()
    await settle()
    const box = (k: string) => document.querySelector<HTMLInputElement>(`[data-xs-protect-box="${k}"]`)!
    expect([box('selectLockedCells').checked, box('formatCells').checked]).toEqual([true, false])
    box('formatCells').click()
    // Select locked cells takes Select unlocked cells along, as in Excel
    box('selectUnlockedCells').click()
    await settle()
    expect(box('selectLockedCells').checked).toBe(false)
    const contents = document.querySelector<HTMLInputElement>('[data-xs-protect-contents]')!
    contents.click()
    await settle()
    expect(document.querySelector<HTMLButtonElement>('[data-xs-protect-ok]')!.disabled).toBe(true)
    contents.click()
    await settle()
    document.querySelector<HTMLButtonElement>('[data-xs-protect-ok]')!.click()
    await until(() => !!lastState(w)?.protection)
    expect((await source.sheets())[0]!.protection).toEqual({ selectLockedCells: false, selectUnlockedCells: false, formatCells: true })
    expect(api(w).protectSheet()).toBe(false)
    await w.find('.xs-tab').trigger('contextmenu', { button: 2, clientX: 20, clientY: 400 })
    await settle()
    menuItem('unprotectSheet')!.click()
    await until(() => !lastState(w)?.protection)
    expect((await source.sheets())[0]!.protection).toBeUndefined()
    expect(api(w).protectSheet()).toBe(true)
    await settle()
    expect(document.querySelector('[data-xs-protect-dialog]')).not.toBeNull()
  })
})
