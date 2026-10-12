// The cell's own text edited in place (F2, a double click; the formula bar opens the same way) goes to the source with
// `inPlace`, a key typed over the cell without it, as Excel's percent editor tells them apart.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource } from '../src'
import { editorKey, goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

describe('an edit in place', () => {
  it('F2 and a double click send inPlace; typing over the cell does not', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: { value: 0.05, text: '5%' } } }] })
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    const last = () => edit.mock.calls.at(-1)
    await goTo(w, 'A1')
    await key(w, 'F2')
    await type(w, '6%')
    await editorKey(w, 'Enter')
    await settle()
    expect(last()).toEqual([[{ sheet: 'S', cell: 'A1', text: '6%' }], { inPlace: true }])
    await goTo(w, 'A1')
    await w.find('[data-cell="1:1"]').trigger('dblclick')
    await type(w, '7%')
    await editorKey(w, 'Enter')
    await settle()
    expect(last()).toEqual([[{ sheet: 'S', cell: 'A1', text: '7%' }], { inPlace: true }])
    await goTo(w, 'A1')
    await key(w, '8')
    await type(w, '8')
    await editorKey(w, 'Enter')
    await settle()
    expect(last()).toEqual([[{ sheet: 'S', cell: 'A1', text: '8' }]])
  })
})
