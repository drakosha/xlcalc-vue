// The active cell's formatting in the `state` event, for a toolbar: its displayed style and its number format code, from
// the window and kept current by edits.
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource } from '../src'
import type { SheetViewState } from '../src'
import type { SheetWindow, WindowRequest } from '../src/contract/source'
import { editorKey, goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

/** An in-memory source whose windows return number formats, as a source over xlcalc does. */
function formatted() {
  const source = memorySource({
    styles: [{ font: { bold: true } }],
    sheets: [{ name: 'S', cells: { A1: { value: 0.05, style: 0 }, B1: 46096, C1: 'text' } }],
  })
  const window = source.window.bind(source)
  const formats: Record<string, string> = { A1: '0%', B1: 'dd.mm.yyyy' }
  source.window = async (req: WindowRequest, signal?: AbortSignal) => {
    const w = (await window(req, signal)) as SheetWindow
    return { ...w, numberFormats: { ...formats } }
  }
  return source
}
const last = (states: (SheetViewState | null)[]) => states[states.length - 1]

describe("state: the active cell's formatting", () => {
  it("returns the active cell's displayed style and its number format code", async () => {
    const w = await mountSheet(formatted())
    const states = () => w.emitted('state')!.map((e) => e[0] as SheetViewState | null)
    await goTo(w, 'A1')
    expect(last(states())).toMatchObject({ ref: 'A1', style: { font: { bold: true } }, numberFormat: '0%' })
    await goTo(w, 'B1')
    expect(last(states())).toMatchObject({ ref: 'B1', numberFormat: 'dd.mm.yyyy' })
    expect(last(states())!.style).toBeUndefined()
    await goTo(w, 'C1')
    expect(last(states())!.numberFormat).toBeUndefined()
  })

  it("takes an edited cell's number format from the outcome", async () => {
    const source = formatted()
    source.edit = async (entries) => ({ ok: true, changes: entries.map((e) => ({ sheet: e.sheet, cell: e.cell, text: '15.03.2026', value: 46096, formula: null, numberFormat: 'dd.mm.yyyy' })) })
    const w = await mountSheet(source)
    await goTo(w, 'C1')
    await key(w, '1')
    await type(w, '15.03.2026')
    await editorKey(w, 'Enter')
    await settle()
    await goTo(w, 'C1')
    const states = w.emitted('state')!.map((e) => e[0] as SheetViewState | null)
    expect(last(states)).toMatchObject({ ref: 'C1', numberFormat: 'dd.mm.yyyy' })
  })
})
