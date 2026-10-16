// References of a formula as the sheet shows them (xlcalc's scanReferences): colours in the bar, ranges on the grid, places to go to.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { sheetContextKey } from '../src/components/context'
import type { SheetContext } from '../src/core/sheet'
import { address, barField, goTo, key, mountSheet, pressBar, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const refs = (w: VueWrapper) => w.findAll('.xs-fx-layer .xs-ref').map((r) => r.text())
const colorOf = (w: VueWrapper, text: string) => (w.findAll('.xs-fx-layer .xs-ref').find((r) => r.text() === text)!.element as HTMLElement).style.color
const tinted = (w: VueWrapper, cell: string) => (w.find(`[data-cell="${cell}"]`).element as HTMLElement).style.backgroundImage.includes('linear-gradient')

describe('a formula in en-US', () => {
  const formula = `SUM(A1:A2)+'Two sheet'!B2+Sales[Qty]+LEN("see C9")+SUM(Jan:Mar!D4)+$A$1+LOG10(A1)`
  const book = (): MemoryBook => ({
    sheets: [
      { name: 'S', cells: { A1: 1, A2: 2, B1: { value: 3, formula }, C9: 'x' } },
      { name: 'Two sheet', cells: { B2: 5 } },
    ],
  })

  it('colours its references to cells in the bar and leaves the rest plain; equal references share a colour', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'B1')
    // A table's reference, one across sheets and a string's text stay text.
    expect(refs(w)).toEqual(['A1:A2', "'Two sheet'!B2", '$A$1', 'A1'])
    expect(barField(w).value).toBe(`=${formula}`)
    expect(colorOf(w, '$A$1')).toBe(colorOf(w, 'A1'))
    expect(colorOf(w, 'A1')).not.toBe(colorOf(w, 'A1:A2'))
    expect(tinted(w, '1:1')).toBe(true)
    expect(tinted(w, '2:1')).toBe(true)
    expect(tinted(w, '9:3')).toBe(false)
  })

  it('highlights a draft as it is typed, a half-typed one too', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'D1')
    await key(w, '=')
    await type(w, '=SUM(C9:C9,A2')
    await settle()
    expect(tinted(w, '9:3')).toBe(true)
    expect(tinted(w, '2:1')).toBe(true)
    expect(tinted(w, '1:1')).toBe(false)
  })
})

describe('the formulas of a selection', () => {
  it("highlight their ranges after the active cell's, each cell's in order; formulas outside it are not read", async () => {
    const book: MemoryBook = { sheets: [{ name: 'S', cells: { B2: { value: 1, formula: 'D5' }, C2: { value: 1, formula: 'G1' }, C3: { value: 1, formula: 'E6+D5' }, B5: { value: 1, formula: 'F1' } } }] }
    const w = await mountSheet(memorySource(book))
    const ctx = (w.vm.$ as unknown as { provides: Record<symbol, SheetContext> }).provides[sheetContextKey as symbol]!
    const ranges = () => [...ctx.f.refColors.value.values()].map((x) => x.ref.text)
    ctx.selected.value = { row: 3, col: 3 }
    ctx.extent.value = { row: 2, col: 2 }
    await settle()
    expect(ranges()).toEqual(['E6', 'D5', 'G1'])
    expect(tinted(w, '6:5') && tinted(w, '5:4') && tinted(w, '1:7')).toBe(true)
    expect(tinted(w, '1:6')).toBe(false)
    ctx.selected.value = { row: 2, col: 2 }
    ctx.extent.value = { row: 3, col: 3 }
    await settle()
    expect(ranges()).toEqual(['D5', 'G1', 'E6'])
  })
})

describe('a formula in ru-RU', () => {
  // An array constant's rows (`2:3`) and decimal commas are no references in ru-RU's formula text.
  const formula = 'СУММ({1;2:3;4};A1)+ЕСЛИ(B2>0,5;C2;D2)'
  const book = (): MemoryBook => ({ sheets: [{ name: 'Лист', cells: { A1: 1, B2: 2, F1: { value: 3, formula }, E3: 4 } }] })

  it("is read in the locale's formula text once its module loads", async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true, locale: 'ru-RU' })
    await goTo(w, 'F1')
    await vi.waitFor(() => expect(refs(w)).toEqual(['A1', 'B2', 'C2', 'D2']))
    expect(barField(w).value).toBe(`=${formula}`)
    expect(tinted(w, '2:3')).toBe(true)
    expect(tinted(w, '3:5')).toBe(false)
  })
})

describe('places', () => {
  const book = (): MemoryBook => ({
    sheets: [
      { name: 'S', cells: { A1: 1, B1: { value: 1, formula: 'STRASSE!A1' }, D5: 1 } },
      { name: 'Straße', cells: { A1: 1 } },
      { name: '2024', cells: { B2: 'y' } },
    ],
  })

  it('the address box takes a range in any case, a sheet quoted or not, a sheet alone; not a formula or a table', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true, messages: { addressBad: 'Bad address' } })
    await goTo(w, 'b2:c3')
    expect(address(w)).toBe('B2:C3')
    await goTo(w, "'2024'!B2")
    expect(w.find('.xs-tab-on').text()).toBe('2024')
    await goTo(w, 'Straße!A1')
    expect(w.find('.xs-tab-on').text()).toBe('Straße')
    await goTo(w, "'S'")
    expect(w.find('.xs-tab-on').text()).toBe('S')
    for (const bad of ['B12+C1', 'Sales[Qty]', 'Jan:Mar!A1']) {
      await goTo(w, bad)
      expect(w.find('.xs-message').text()).toBe('Bad address')
    }
  })

  it('a sheet in a reference is found as the book compares names; the link quotes a sheet as Excel writes it', async () => {
    const writeText = vi.fn(async () => {})
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
    const w = await mountSheet(memorySource(book()), { readOnly: true, cellUrl: (ref: string) => `at=${ref}` })
    await goTo(w, 'B1')
    await pressBar(w, 3, 3, { ctrlKey: true })
    expect(w.find('.xs-tab-on').text()).toBe('Straße')
    await goTo(w, "'2024'!B2")
    await w.find('.xs-tools .xs-tool').trigger('click')
    await settle()
    expect(writeText).toHaveBeenCalledWith("at='2024'!B2")
  })
})
