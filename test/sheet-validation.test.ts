// Data validation on the sheet: a rejected value shows Excel's alert, whose answer writes it with `accept`, reopens the
// editor or leaves the cell unchanged; a list's arrow and Alt+Down enter an item; the input message appears next to the cell.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { active, cellText, editorKey, editorOpen, goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 5 }, B1: { value: 'да' } },
      validations: [
        { range: 'A1', whole: { min: 1, max: 10 }, alert: { severity: 'stop', title: 'Число', message: 'От 1 до 10' } },
        { range: 'A2', whole: { min: 1, max: 10 }, alert: { severity: 'warning', message: 'Лучше от 1 до 10' } },
        { range: 'B1', list: ['да', 'нет', 'может быть'], prompt: { title: 'Ответ', text: 'Выберите из списка' } },
      ],
    },
  ],
})
const alertShown = () => document.querySelector('[data-xs-alert]')
const press = async (button: string) => {
  document.querySelector<HTMLElement>(`[data-xs-alert-button="${button}"]`)!.click()
  await settle()
}

describe('data validation on the sheet', () => {
  it('a stop alert: Retry reopens the editor with the typed text, Cancel leaves the active cell unchanged', async () => {
    const source = memorySource(book())
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await goTo(w, 'A1')
    await key(w, '5')
    await type(w, '50')
    await editorKey(w, 'Enter')
    await settle()
    expect(alertShown()?.textContent).toContain('От 1 до 10')
    expect(active(w)).toBe('1:1')
    await press('retry')
    expect(alertShown()).toBeNull()
    expect(editorOpen(w)).toBe(true)
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('50')
    await editorKey(w, 'Enter')
    await settle()
    expect(alertShown()).not.toBeNull()
    await press('cancel')
    expect(alertShown()).toBeNull()
    expect(cellText(w, 1, 1)).toBe('5')
    expect(active(w)).toBe('1:1')
    expect(edit.mock.calls.every((c) => c[1] === undefined)).toBe(true)
  })

  it('a warning alert: Yes writes the value with accept', async () => {
    const source = memorySource(book())
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await goTo(w, 'A2')
    await key(w, '5')
    await type(w, '50')
    await editorKey(w, 'Enter')
    await settle()
    expect(alertShown()?.textContent).toContain('Лучше от 1 до 10')
    await press('yes')
    await settle()
    expect(edit.mock.calls.at(-1)).toEqual([[{ sheet: 'S', cell: 'A2', text: '50' }], { accept: true }])
    expect(cellText(w, 2, 1)).toBe('50')
  })

  it('a list: the arrow next to the active cell, Alt+Down opens the items, Enter enters one past the rule; the input message appears', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    const source = memorySource(book())
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await goTo(w, 'B1')
    await vi.advanceTimersByTimeAsync(100)
    await settle()
    expect(w.find('[data-xs-dropdown-arrow]').exists()).toBe(true)
    expect(w.find('[data-xs-prompt]').text()).toContain('Выберите из списка')
    await key(w, 'ArrowDown', { altKey: true })
    await vi.advanceTimersByTimeAsync(100)
    await settle()
    expect(w.findAll('[data-xs-dropdown-item]').map((i) => i.text())).toEqual(['да', 'нет', 'может быть'])
    await key(w, 'ArrowDown')
    await key(w, 'Enter')
    await settle()
    expect(edit.mock.calls.at(-1)).toEqual([[{ sheet: 'S', cell: 'B1', text: 'нет' }], { accept: true }])
    expect(cellText(w, 1, 2)).toBe('нет')
    await goTo(w, 'A1')
    await vi.advanceTimersByTimeAsync(100)
    await settle()
    expect(w.find('[data-xs-dropdown-arrow]').exists()).toBe(false)
    expect(w.find('[data-xs-prompt]').exists()).toBe(false)
  })

  it("a pick the source rejects requests the cell's rule again, so the list shows what the source holds now", async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    const source = memorySource(book())
    const validation = vi.spyOn(source, 'validation')
    const pick = vi.fn(async () => ({ ok: false as const, reason: 'conflict' as const, message: 'The list changed' }))
    const w = await mountSheet({ ...source, pick })
    await goTo(w, 'B1')
    await vi.advanceTimersByTimeAsync(100)
    await settle()
    await key(w, 'ArrowDown', { altKey: true })
    await vi.advanceTimersByTimeAsync(100)
    await settle()
    const asked = validation.mock.calls.length
    await key(w, 'Enter')
    await vi.advanceTimersByTimeAsync(100)
    await settle()
    expect(pick).toHaveBeenCalledTimes(1)
    expect(validation.mock.calls.length).toBeGreaterThan(asked)
  })

  it('a list over a source with pick: an item goes to pick by its position and text, not to edit', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    const source = memorySource(book())
    const edit = vi.spyOn(source, 'edit')
    const pick = vi.fn(async (sheet: string, cell: string, _item: number, text: string) => ({ ok: true as const, changes: [{ sheet, cell, text, value: text }] }))
    const w = await mountSheet({ ...source, pick })
    await goTo(w, 'B1')
    await vi.advanceTimersByTimeAsync(100)
    await settle()
    await key(w, 'ArrowDown', { altKey: true })
    await vi.advanceTimersByTimeAsync(100)
    await settle()
    await key(w, 'ArrowDown')
    await key(w, 'ArrowDown')
    await key(w, 'Enter')
    await settle()
    expect(pick.mock.calls).toEqual([['S', 'B1', 2, 'может быть']])
    expect(edit).not.toHaveBeenCalled()
    expect(cellText(w, 1, 2)).toBe('может быть')
  })
})
