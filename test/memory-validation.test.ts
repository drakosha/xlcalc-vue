// memorySource and data validation: a sheet's simple rules (a list, whole numbers within bounds), a cell's dropdown and
// input message, a value refused with its alert unless the edit is accepted.
import { describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook } from '../src'

const book = (): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: 'Да', B1: 5 },
      validations: [
        { range: 'A1:A9 C1', list: ['Да', 'Нет'], prompt: { title: 'Ответ', text: 'Да или нет' } },
        { range: 'B:B', whole: { min: 1, max: 10 }, alert: { title: 'Число', message: 'От 1 до 10' } },
        { range: 'D1:D9', whole: { min: 0 }, alert: { severity: 'warning' } },
        { range: 'E1:E9', list: ['x'], alert: { severity: 'information', message: 'Не из списка' } },
      ],
    },
    { name: 'T', cells: {} },
  ],
})

describe('memorySource: data validation', () => {
  it('gives a cell’s list and input message, a rule without them empty, no rule null', async () => {
    const src = memorySource(book())
    expect(await src.validation!('S', 'A5')).toEqual({ list: ['Да', 'Нет'], prompt: { title: 'Ответ', text: 'Да или нет' } })
    expect(await src.validation!('S', 'C1')).toEqual({ list: ['Да', 'Нет'], prompt: { title: 'Ответ', text: 'Да или нет' } })
    expect(await src.validation!('S', 'B900')).toEqual({})
    expect(await src.validation!('S', 'C2')).toBeNull()
    expect(await src.validation!('T', 'A1')).toBeNull()
  })

  it('refuses a value off the list or out of bounds with the rule’s alert, Excel’s title and text by default', async () => {
    const src = memorySource(book())
    expect(await src.edit!([{ sheet: 'S', cell: 'B2', text: '11' }])).toEqual({
      ok: false,
      reason: 'refused',
      message: 'От 1 до 10',
      alert: { severity: 'stop', title: 'Число', message: 'От 1 до 10', sheet: 'S', cell: 'B2' },
    })
    expect(await src.edit!([{ sheet: 'S', cell: 'B2', text: '2.5' }])).toMatchObject({ ok: false, alert: { severity: 'stop' } })
    expect(await src.edit!([{ sheet: 'S', cell: 'B2', text: 'пять' }])).toMatchObject({ ok: false, alert: { severity: 'stop' } })
    expect(await src.edit!([{ sheet: 'S', cell: 'A2', text: 'да' }])).toEqual({
      ok: false,
      reason: 'refused',
      message: "This value doesn't match the data validation restrictions defined for this cell.",
      alert: { severity: 'stop', message: "This value doesn't match the data validation restrictions defined for this cell.", sheet: 'S', cell: 'A2' },
    })
    expect(await src.edit!([{ sheet: 'S', cell: 'D2', text: '-1' }])).toMatchObject({ ok: false, alert: { severity: 'warning', cell: 'D2' } })
    expect(await src.edit!([{ sheet: 'S', cell: 'E2', text: 'y' }])).toMatchObject({ ok: false, alert: { severity: 'information', message: 'Не из списка' } })
    const w = await src.window({ sheet: 'S', row: 2, rows: 1, col: 1, cols: 5 })
    expect('text' in w && w.text).toEqual([['', '', '', '', '']])
  })

  it('takes what the rule takes: an item with spaces about it, a whole number within, a blank, a formula', async () => {
    const src = memorySource(book())
    expect(await src.edit!([{ sheet: 'S', cell: 'A2', text: ' Нет ' }, { sheet: 'S', cell: 'B2', text: '10' }, { sheet: 'S', cell: 'B3', text: '=B1*9' }, { sheet: 'S', cell: 'B1', text: null }])).toMatchObject({ ok: true })
  })

  it('writes a refused value with `accept`; a stop rule of several cells refuses before a warning', async () => {
    const src = memorySource(book())
    expect(await src.edit!([{ sheet: 'S', cell: 'D2', text: '-1' }, { sheet: 'S', cell: 'B2', text: '0' }])).toMatchObject({ ok: false, alert: { severity: 'stop', cell: 'B2' } })
    expect(await src.edit!([{ sheet: 'S', cell: 'D2', text: '-1' }], { accept: true })).toMatchObject({ ok: true, changes: [{ cell: 'D2', value: -1 }] })
  })
})
