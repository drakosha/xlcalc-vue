// Excel's calculation options on the sheet: in Manual the status bar shows Calculate while cells wait, F9 calculates the
// workbook and Shift+F9 the active sheet, the state reports the option, and the host sets it through the sheet.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { BookInput } from 'xlcalc'
import { memorySource, type SheetViewApi, type SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { cellText, editorKey, goTo, key, mountSheet, type, until } from './helpers'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

const book = (calculation?: BookInput['calculation']): BookInput => ({
  ...(calculation ? { calculation } : {}),
  sheets: [
    { name: 'One', cells: { A1: { value: 1 }, B1: { formula: 'A1*2' } } },
    { name: 'Two', cells: { A1: { formula: 'One!A1*100' } } },
  ],
})
function open(calculation?: BookInput['calculation']): LocalSource {
  const s = localSource(book(calculation), { calc: 'inline' })
  sources.push(s)
  return s
}
const status = (w: VueWrapper) => w.find('[data-xs-calculate]')
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
async function enter(w: VueWrapper, place: string, text: string): Promise<void> {
  await goTo(w, place)
  await key(w, 'F2')
  await type(w, text)
  await editorKey(w, 'Enter')
}

describe('calculation options on the sheet', () => {
  it('in Manual an edit leaves the dependent as it was, the status bar shows Calculate, and F9 calculates', async () => {
    const w = await mountSheet(open('manual'))
    await until(() => cellText(w, 1, 2) === '2')
    expect(lastState(w)).toMatchObject({ calculation: 'manual' })
    expect(status(w).exists()).toBe(false)
    await enter(w, 'A1', '5')
    await until(() => status(w).exists())
    expect(cellText(w, 1, 1)).toBe('5')
    expect(cellText(w, 1, 2)).toBe('2')
    expect(status(w).text()).toBe('Calculate')
    expect(lastState(w)).toMatchObject({ calculation: 'manual', calculationPending: true })
    await key(w, 'F9')
    await until(() => cellText(w, 1, 2) === '10')
    await until(() => !status(w).exists())
    expect(lastState(w)?.calculationPending).toBeUndefined()
  })

  it('Shift+F9 calculates the active sheet only; a click on Calculate calculates the workbook', async () => {
    const s = open('manual')
    const calculate = vi.spyOn(s, 'calculate')
    const w = await mountSheet(s)
    await until(() => cellText(w, 1, 2) === '2')
    await enter(w, 'A1', '5')
    await until(() => status(w).exists())
    await key(w, 'F9', { shiftKey: true })
    await until(() => cellText(w, 1, 2) === '10')
    expect(calculate).toHaveBeenLastCalledWith('One')
    // sheet Two still waits
    expect(status(w).exists()).toBe(true)
    await status(w).trigger('click')
    await until(() => !status(w).exists())
    expect(calculate).toHaveBeenLastCalledWith()
  })

  it('the host sets the option through the sheet; a read-only sheet sends nothing', async () => {
    const s = open()
    const w = await mountSheet(s)
    await until(() => lastState(w)?.calculation === 'automatic')
    const api = w.vm as unknown as SheetViewApi
    expect(await api.setCalculation('manual')).toEqual({ ok: true, changes: [] })
    await until(() => lastState(w)?.calculation === 'manual')
    await w.setProps({ readOnly: true })
    expect(await api.setCalculation('automatic')).toBeNull()
    expect(await api.calculate()).toBeNull()
  })

  it('a source without calculation reports none and F9 does nothing', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }))
    expect(lastState(w)?.calculation).toBeUndefined()
    await key(w, 'F9')
    expect(await (w.vm as unknown as SheetViewApi).calculate()).toBeNull()
  })
})
