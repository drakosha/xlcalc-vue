// Excel's AutoSum (Alt+=): the source's guess opens the active cell's edit with its SUM and the range selected; `./local`
// guesses through xlcalc's getAutoSum, with the formula in the workbook's locale.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type SheetViewApi } from '../src'
import { localSource } from '../src/local'
import { editorKey, editorOpen, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const numbers = { sheets: [{ name: 'S', cells: { B2: 1, B3: 2, B4: 3 } }] }
const field = (w: Awaited<ReturnType<typeof mountSheet>>) => w.find('.xs-editor').element as HTMLTextAreaElement
const altEquals = (w: Awaited<ReturnType<typeof mountSheet>>) => key(w, '≠', { code: 'Equal', altKey: true })

describe('AutoSum in ./local', () => {
  it("guesses the run of numbers above the cell, in the workbook's locale, and SUM() without numbers", async () => {
    const en = localSource({ sheets: [{ name: 'S', cells: { B2: { value: 1 }, B3: { value: 2 }, B4: { value: 3 } } }] }, { calc: 'inline' })
    expect(await en.autoSum('S', 'B5')).toEqual({ range: 'B2:B4', formula: 'SUM(B2:B4)' })
    expect(await en.autoSum('S', 'D1')).toEqual({ range: null, formula: 'SUM()' })
    await en.close()
    const ru = localSource({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 1 }, A2: { value: 2 } } }] }, { calc: 'inline' })
    expect(await ru.autoSum('Лист1', 'A3')).toEqual({ range: 'A1:A2', formula: 'СУММ(A1:A2)' })
    await ru.close()
  })
})

describe('AutoSum on the sheet', () => {
  it('Alt+= opens the edit with the SUM and its range selected; Enter enters it', async () => {
    const source = memorySource(numbers)
    source.autoSum = vi.fn(async () => ({ range: 'B2:B4', formula: 'SUM(B2:B4)' }))
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await goTo(w, 'B5')
    await altEquals(w)
    await settle()
    expect(source.autoSum).toHaveBeenCalledWith('S', 'B5')
    expect(editorOpen(w)).toBe(true)
    expect(field(w).value).toBe('=SUM(B2:B4)')
    expect([field(w).selectionStart, field(w).selectionEnd]).toEqual([5, 10])
    await editorKey(w, 'Enter')
    await settle()
    expect(edit.mock.calls[0]![0]).toEqual([{ sheet: 'S', cell: 'B5', text: '=SUM(B2:B4)' }])
  })

  it('without a range the caret is between the parentheses; the method does the same', async () => {
    const source = memorySource(numbers)
    source.autoSum = async () => ({ range: null, formula: 'SUM()' })
    const w = await mountSheet(source)
    await goTo(w, 'D1')
    expect(await (w.vm as unknown as SheetViewApi).autoSum()).toBe(true)
    await settle()
    expect(field(w).value).toBe('=SUM()')
    expect([field(w).selectionStart, field(w).selectionEnd]).toEqual([5, 5])
  })

  it('without the source’s autoSum, or on a read-only sheet, there is none', async () => {
    const w = await mountSheet(memorySource(numbers))
    expect(await (w.vm as unknown as SheetViewApi).autoSum()).toBe(false)
    const source = memorySource(numbers)
    source.autoSum = vi.fn(async () => ({ range: null, formula: 'SUM()' }))
    const ro = await mountSheet(source, { readOnly: true })
    await altEquals(ro)
    await settle()
    expect(source.autoSum).not.toHaveBeenCalled()
    expect(ro.find('.xs-editor').exists()).toBe(false)
  })
})
