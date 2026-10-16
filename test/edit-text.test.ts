// The text a cell opens with when its source gives no `editText`: xlcalc's `entryText` in the book's locale by the cell's
// number format (`xlcalc/entry`, a chunk loaded on the first such cell), the shown text until the chunk is in.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryCell } from '../src'
import * as entry from '../src/entry'
import { barField, editorKey, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

/** A memory book whose windows carry `numberFormats`, as a source over a real book gives them. */
function withFormats(cells: Record<string, MemoryCell>, formats: Record<string, string>) {
  const source = memorySource({ sheets: [{ name: 'S', cells }] })
  const window = source.window.bind(source)
  source.window = async (req, signal) => {
    const got = await window(req, signal)
    return 'pending' in got ? got : { ...got, numberFormats: { ...formats } }
  }
  return source
}
/** What the field opens with on `place` by F2. */
async function open(w: VueWrapper, place: string): Promise<string> {
  await goTo(w, place)
  await key(w, 'F2')
  await settle()
  const text = (w.find('.xs-editor').element as HTMLTextAreaElement).value
  await editorKey(w, 'Escape')
  return text
}

const CELLS: Record<string, MemoryCell> = {
  A1: { value: 1234.5678, text: '1 234,57 ₽' },
  A2: { value: 0.355, text: '35,5%' },
  A3: { value: 46096, text: '15.03.2026' },
  A4: { value: 0.75, text: '18:00' },
  A5: { value: '007' },
  A6: { value: '1,5' },
  A7: { value: true, text: 'ИСТИНА' },
  A8: { value: 'Отдел 5' },
}
const FORMATS = { A1: '#,##0.00 ₽', A2: '0.0%', A3: 'dd.mm.yyyy', A4: 'h:mm' }

describe('the text a cell opens with, by xlcalc/entry', () => {
  it("in ru-RU: the full number with the decimal comma, a percent, a date, a time as Excel's bar writes them, Excel's apostrophe and words", async () => {
    const w = await mountSheet(withFormats(CELLS, FORMATS), { locale: 'ru-RU' })
    expect(await open(w, 'A1')).toBe('1234,5678')
    expect(await open(w, 'A2')).toBe('35,5%')
    expect(await open(w, 'A3')).toBe('15.03.2026')
    expect(await open(w, 'A4')).toBe('18:00:00')
    expect(await open(w, 'A5')).toBe("'007")
    expect(await open(w, 'A6')).toBe("'1,5")
    expect(await open(w, 'A7')).toBe('ИСТИНА')
    expect(await open(w, 'A8')).toBe('Отдел 5')
    expect(barField(w).value).toBe('Отдел 5')
  })

  it('in en-US: the apostrophe goes where en-US would read a value, `1,5` is text there', async () => {
    const w = await mountSheet(withFormats(CELLS, FORMATS))
    expect(await open(w, 'A1')).toBe('1234.5678')
    expect(await open(w, 'A5')).toBe("'007")
    expect(await open(w, 'A6')).toBe('1,5')
    expect(await open(w, 'A7')).toBe('TRUE')
  })

  it('without number formats a number opens whole, as General', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: CELLS }] }), { locale: 'ru-RU' })
    expect(await open(w, 'A2')).toBe('0,355')
  })

  it('opens with the shown text at once while the chunk loads, and the formula bar takes the entry text once it is in', async () => {
    let arrive: () => void = () => {}
    const real = entry.loadEntry
    const gate = new Promise<void>((r) => (arrive = r))
    const load = vi.spyOn(entry, 'loadEntry').mockImplementation(() => gate.then(real))
    const w = await mountSheet(withFormats(CELLS, FORMATS), { locale: 'ru-RU' })
    expect(await open(w, 'A1')).toBe('1 234,57 ₽')
    await goTo(w, 'A2')
    expect(barField(w).value).toBe('35,5%')
    arrive()
    await settle()
    await goTo(w, 'A1')
    expect(barField(w).value).toBe('1234,5678')
    expect(load).toHaveBeenCalledTimes(1)
  })
})
