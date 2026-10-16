// @vitest-environment node
// localSource and data validation: a cell's dropdown and input message, a value refused with Excel's alert for any style,
// written after the person keeps it (`accept`), a paste not checked, as Excel.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

/** A ru-RU sheet with rules of each kind: lists literal and of cells, whole numbers of the three styles, a message alone. */
const book = (): BookInput => ({
  locale: 'ru-RU',
  sheets: [
    {
      name: 'S',
      cells: { E1: { value: 1.5, numberFormat: '0.00' }, E2: { value: 'два' }, E4: { value: 'четыре' }, A1: { value: 'Да' }, B1: { value: 5 } },
      validations: [
        { range: 'A1:A9', type: 'list', formula1: '"Да,Нет,Может быть"', promptTitle: 'Ответ', prompt: 'Выберите из списка' },
        { range: 'B1:B9', type: 'whole', operator: 'between', formula1: '1', formula2: '10', errorStyle: 'stop', errorTitle: 'Число', error: 'От 1 до 10' },
        { range: 'C1:C9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning', error: 'Лучше больше нуля' },
        { range: 'D1:D9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'information' },
        { range: 'F1:F9', type: 'list', formula1: '$E$1:$E$9' },
        { range: 'G1', type: 'list', formula1: 'Нет_такого_имени' },
        { range: 'H1', type: 'list', formula1: '"x,y"', dropdown: false, showPrompt: false, prompt: 'не показывается' },
        { range: 'I1', type: 'any', prompt: 'Только текст подсказки' },
        { range: 'J1', type: 'whole', operator: 'greaterThan', formula1: '0', showError: false },
      ],
    },
  ],
})

describe('localSource: a cell’s dropdown and input message', () => {
  it('gives a literal list’s items and the input message, a list of cells as their cells show, blanks inside kept', async () => {
    const s = source(book())
    expect(await s.validation!('S', 'A3')).toEqual({ list: ['Да', 'Нет', 'Может быть'], prompt: { title: 'Ответ', text: 'Выберите из списка' } })
    expect(await s.validation!('S', 'F2')).toEqual({ list: ['1,50', 'два', '', 'четыре'] })
  })

  it('gives an empty list for a formula that gives none, no list without the dropdown, no message turned off', async () => {
    const s = source(book())
    expect(await s.validation!('S', 'G1')).toEqual({ list: [] })
    expect(await s.validation!('S', 'H1')).toEqual({})
    expect(await s.validation!('S', 'I1')).toEqual({ prompt: { text: 'Только текст подсказки' } })
    expect(await s.validation!('S', 'B2')).toEqual({})
    expect(await s.validation!('S', 'K1')).toBeNull()
  })

  it('reads the list as the book is now', async () => {
    const s = source(book())
    await s.edit([{ sheet: 'S', cell: 'E3', text: 'три' }])
    expect((await s.validation!('S', 'F1'))?.list).toEqual(['1,50', 'два', 'три', 'четыре'])
  })

  it('stops with an aborted signal', async () => {
    const s = source(book())
    const ctrl = new AbortController()
    ctrl.abort()
    await expect(s.validation!('S', 'A1', ctrl.signal)).rejects.toMatchObject({ name: 'AbortError' })
  })
})

describe('localSource: a value a rule does not take', () => {
  it('refuses it for any style with the alert, nothing written', async () => {
    const s = source(book())
    expect(await s.edit([{ sheet: 'S', cell: 'B2', text: '11' }])).toEqual({
      ok: false,
      reason: 'refused',
      message: 'От 1 до 10',
      alert: { severity: 'stop', title: 'Число', message: 'От 1 до 10', sheet: 'S', cell: 'B2' },
    })
    expect(await s.edit([{ sheet: 'S', cell: 'C2', text: '-1' }])).toMatchObject({ ok: false, reason: 'refused', alert: { severity: 'warning', message: 'Лучше больше нуля', sheet: 'S', cell: 'C2' } })
    const info = await s.edit([{ sheet: 'S', cell: 'D2', text: '-2' }])
    // without its own title and message: the component's title, Excel's text in the book's locale (xlcalc 0.28)
    expect(info).toMatchObject({ ok: false, alert: { severity: 'information', message: 'Это значение не соответствует ограничениям по проверке данных, установленным для этой ячейки.' } })
    expect(!info.ok && info.alert && 'title' in info.alert).toBe(false)
    expect(await s.edit([{ sheet: 'S', cell: 'A2', text: 'да ' }])).toMatchObject({ ok: false, alert: { severity: 'stop', cell: 'A2' } })
    expect((await win(s, 'S', 2, 1, 1, 4)).values).toEqual([[null, null, null, null]])
    expect(s.history).toEqual({ undo: 0, redo: 0 })
  })

  it('writes it with `accept`, as Excel after Yes or OK, in one step of the undo', async () => {
    const s = source(book())
    expect(await s.edit([{ sheet: 'S', cell: 'C2', text: '-1' }], { accept: true })).toMatchObject({ ok: true, changes: [{ cell: 'C2', value: -1 }] })
    expect(await s.edit([{ sheet: 'S', cell: 'D2', text: '-2' }], { accept: true })).toMatchObject({ ok: true, changes: [{ cell: 'D2', value: -2 }] })
    expect(s.history).toEqual({ undo: 2, redo: 0 })
    // a formula that does not read is refused all the same
    expect(await s.edit([{ sheet: 'S', cell: 'C3', text: '=СУММ(' }], { accept: true })).toMatchObject({ ok: false, reason: 'invalid' })
  })

  it('refuses several cells by the strictest rule they break, the warning before the stop given first', async () => {
    const s = source(book())
    const got = await s.edit([{ sheet: 'S', cell: 'C2', text: '-1' }, { sheet: 'S', cell: 'D2', text: '-1' }, { sheet: 'S', cell: 'B3', text: '0' }])
    expect(got).toMatchObject({ ok: false, alert: { severity: 'stop', cell: 'B3' } })
    expect(await s.edit([{ sheet: 'S', cell: 'D2', text: '-1' }, { sheet: 'S', cell: 'C2', text: '-1' }])).toMatchObject({ ok: false, alert: { severity: 'warning', cell: 'C2' } })
  })

  it('takes any value under a rule that shows no alert, and a cleared cell', async () => {
    const s = source(book())
    expect(await s.edit([{ sheet: 'S', cell: 'J1', text: '-5' }])).toMatchObject({ ok: true })
    expect(await s.edit([{ sheet: 'S', cell: 'B1', text: null }])).toMatchObject({ ok: true })
  })

  it('Control+Return checks the active cell’s rule, and fills with `accept`', async () => {
    const s = source(book())
    expect(await s.fill!('S', 'C1:C3', '-7', 'C2')).toMatchObject({ ok: false, reason: 'refused', alert: { severity: 'warning', sheet: 'S', cell: 'C2' } })
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    expect(await s.fill!('S', 'C1:C3', '-7', 'C2', { accept: true })).toMatchObject({ ok: true, select: { sheet: 'S', range: 'C1:C3' } })
    expect((await win(s, 'S', 1, 3, 3, 1)).values).toEqual([[-7], [-7], [-7]])
  })

  it('pastes without checking, as Excel', async () => {
    const s = source(book())
    expect(await s.paste!('S', 'B2', '99\t-3')).toMatchObject({ ok: true })
    expect((await win(s, 'S', 2, 1, 2, 2)).values).toEqual([[99, -3]])
  })
})

describe('localSource: an item picked from a dropdown', () => {
  /** ru-RU: a list of cells holding the text 001, the text 1/2 and 0.5 shown as 50 %; a literal list; both over General cells. */
  const lists = (): BookInput => ({
    locale: 'ru-RU',
    sheets: [
      {
        name: 'S',
        cells: { E1: { value: '001' }, E2: { value: '1/2' }, E3: { value: 0.5, numberFormat: '0%' }, A1: { value: 'x' } },
        validations: [
          { range: 'A1:A9', type: 'list', formula1: '$E$1:$E$3' },
          { range: 'B1:B9', type: 'list', formula1: '"001,1/2,x"' },
        ],
      },
    ],
  })

  it('writes a cell\'s item as its value, not as typed, and keeps the cell\'s format, in one step of undo', async () => {
    const s = source(lists())
    expect(await s.validation('S', 'A1')).toEqual({ list: ['001', '1/2', '50%'] })
    expect(await s.pick('S', 'A1', 0, '001')).toMatchObject({ ok: true, changes: [{ cell: 'A1', value: '001', text: '001' }] })
    expect(await s.pick('S', 'A2', 1, '1/2')).toMatchObject({ ok: true, changes: [{ cell: 'A2', value: '1/2' }] })
    expect(await s.pick('S', 'A3', 2, '50%')).toMatchObject({ ok: true, changes: [{ cell: 'A3', value: 0.5, text: '0,5' }] })
    expect((await win(s, 'S', 1, 3, 1, 1)).numberFormats ?? {}).toEqual({})
    expect(s.history).toEqual({ undo: 3, redo: 0 })
    expect(await s.undo()).toMatchObject({ ok: true, select: { sheet: 'S', range: 'A3' } })
    expect((await win(s, 'S', 1, 3, 1, 1)).values).toEqual([['001'], ['1/2'], [null]])
  })

  it('reads a literal item as typed in en-US, whatever the book\'s locale', async () => {
    const s = source(lists())
    expect(await s.pick('S', 'B1', 0, '001')).toMatchObject({ ok: true, changes: [{ cell: 'B1', value: 1 }] })
    expect(await s.pick('S', 'B2', 1, '1/2')).toMatchObject({ ok: true })
    // the 2nd of January of this year in Excel's built-in d-mmm (16), as the book stores it
    const year = new Date().getFullYear()
    expect((await s.book()).sheets[0]!.cells.B2).toEqual({ value: (Date.UTC(year, 0, 2) - Date.UTC(1899, 11, 30)) / 86_400_000, numberFormat: 16 })
  })

  it('refuses an item whose text the list no longer has at its place, nothing written', async () => {
    const s = source(lists())
    expect(await s.pick('S', 'A1', 0, '002')).toMatchObject({ ok: false, reason: 'conflict' })
    expect(await s.pick('S', 'A1', 7, '001')).toMatchObject({ ok: false, reason: 'conflict' })
    expect((await win(s, 'S', 1, 1, 1, 1)).values).toEqual([['x']])
    expect(s.history).toEqual({ undo: 0, redo: 0 })
  })
})
