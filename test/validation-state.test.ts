// The composables of data validation on their own: the active cell's rule asked once the cell stays (an older request
// aborted, its answer dropped), the list's keys and choice, and an alert's answer.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { effectScope, nextTick, ref } from 'vue'
import { ALERT_BUTTONS, ALERT_DEFAULT, useValidation, useValidationAlert, VALIDATION_DELAY_MS, type ValidationDeps } from '../src/core/validation'
import type { SheetSource, SheetValidation, SheetValidationAlert } from '../src/source'

beforeEach(() => vi.useFakeTimers())
afterEach(() => vi.useRealTimers())

/** A source whose `validation` the test answers: each call is kept with its signal. */
function rules(answer: (cell: string) => SheetValidation | null = () => ({ list: ['Да', 'Нет', 'Может быть'], prompt: { text: 'Выберите' } })) {
  const calls: { sheet: string; cell: string; signal?: AbortSignal; resolve: () => void }[] = []
  const source: SheetSource = {
    sheets: async () => [],
    window: async () => ({ pending: { elapsedMs: 0 } }),
    validation: (sheet, cell, signal) => new Promise((resolve) => calls.push({ sheet, cell, signal, resolve: () => resolve(answer(cell)) })),
  }
  return { source, calls }
}
const key = (k: string, more: KeyboardEventInit = {}) => new KeyboardEvent('keydown', { key: k, cancelable: true, ...more })

function setup(source: SheetSource, more: Partial<ValidationDeps> = {}) {
  const active = ref<{ row: number; col: number } | null>({ row: 1, col: 1 })
  const sheet = ref('S')
  const text = ref('')
  const pick = vi.fn()
  const scope = effectScope()
  const v = scope.run(() => useValidation({ source: () => source, sheet: () => sheet.value, active: () => active.value, text: () => text.value, pick, ...more }))!
  return { v, active, sheet, text, pick, scope }
}
/** Let the pause pass and the answers come. */
async function pass(ms = VALIDATION_DELAY_MS): Promise<void> {
  await vi.advanceTimersByTimeAsync(ms)
  await nextTick()
}

describe('useValidation: the active cell’s rule', () => {
  it('asks once the cell stays: arrows over cells ask for the last one alone', async () => {
    const { source, calls } = rules()
    const { v, active } = setup(source)
    await nextTick()
    active.value = { row: 2, col: 1 }
    await nextTick()
    active.value = { row: 3, col: 1 }
    await nextTick()
    expect(calls).toHaveLength(0)
    await pass()
    expect(calls.map((c) => c.cell)).toEqual(['A3'])
    calls[0]!.resolve()
    await pass(0)
    expect(v.current.value).toEqual({ list: ['Да', 'Нет', 'Может быть'], prompt: { text: 'Выберите' } })
    expect(v.place.value).toEqual({ sheet: 'S', cell: 'A3' })
    expect(v.prompt.value).toEqual({ text: 'Выберите' })
  })

  it('aborts the request of a cell left, drops its late answer, and shows nothing meanwhile', async () => {
    const { source, calls } = rules((cell) => (cell === 'B1' ? { prompt: { text: 'B' } } : { prompt: { text: 'A' } }))
    const { v, active } = setup(source)
    await pass()
    calls[0]!.resolve()
    await pass(0)
    expect(v.prompt.value).toEqual({ text: 'A' })
    active.value = { row: 1, col: 2 }
    await nextTick()
    expect(v.current.value).toBeNull()
    await pass()
    active.value = { row: 1, col: 3 }
    await nextTick()
    expect(calls[1]!.signal!.aborted).toBe(true)
    calls[1]!.resolve()
    await pass(0)
    expect(v.current.value).toBeNull()
  })

  it('asks again on another sheet and stops with its signal; a source without `validation` is not asked', async () => {
    const { source, calls } = rules()
    const ctrl = new AbortController()
    const { sheet } = setup(source, { signal: ctrl.signal })
    await pass()
    sheet.value = 'T'
    await pass()
    expect(calls.map((c) => `${c.sheet}!${c.cell}`)).toEqual(['S!A1', 'T!A1'])
    ctrl.abort()
    expect(calls[1]!.signal!.aborted).toBe(true)
    sheet.value = 'U'
    await pass()
    expect(calls).toHaveLength(2)
    const bare: SheetSource = { sheets: async () => [], window: async () => ({ pending: { elapsedMs: 0 } }) }
    const { v } = setup(bare)
    await pass()
    expect(v.current.value).toBeNull()
  })
})

describe('useValidation: the list', () => {
  async function ready() {
    const r = rules()
    const s = setup(r.source)
    await pass()
    r.calls[0]!.resolve()
    await pass(0)
    return { ...s, calls: r.calls }
  }

  it('Alt+Down opens it at the cell’s item; the arrows, Page keys, Home and End move; Enter enters the item and closes', async () => {
    const { v, text, pick, calls } = await ready()
    text.value = 'Нет'
    expect(v.onKey(key('ArrowDown'))).toBe(false)
    const alt = key('ArrowDown', { altKey: true })
    expect(v.onKey(alt)).toBe(true)
    expect(alt.defaultPrevented).toBe(true)
    expect([v.open.value, v.index.value]).toEqual([true, 1])
    // its items are asked again, as Excel computes them on opening
    expect(calls).toHaveLength(2)
    v.onKey(key('ArrowDown'))
    v.onKey(key('ArrowDown'))
    expect(v.index.value).toBe(2)
    v.onKey(key('Home'))
    expect(v.index.value).toBe(0)
    v.onKey(key('PageDown'))
    expect(v.index.value).toBe(2)
    v.onKey(key('ArrowUp'))
    expect(v.onKey(key('Enter'))).toBe(true)
    expect(pick).toHaveBeenCalledWith('Нет', { sheet: 'S', cell: 'A1' })
    expect(v.open.value).toBe(false)
  })

  it('Esc and Alt+Up close it; another key closes it and goes on to the grid; a new cell closes it', async () => {
    const { v, active, pick } = await ready()
    v.openList()
    expect(v.index.value).toBe(-1)
    v.onKey(key('ArrowUp'))
    expect(v.index.value).toBe(2)
    expect(v.onKey(key('Escape'))).toBe(true)
    expect(v.open.value).toBe(false)
    v.toggle()
    v.onKey(key('ArrowUp', { altKey: true }))
    expect(v.open.value).toBe(false)
    v.toggle()
    expect(v.onKey(key('x'))).toBe(false)
    expect(v.open.value).toBe(false)
    v.toggle()
    // Enter on no item only closes
    v.onKey(key('Enter'))
    expect(v.open.value).toBe(false)
    v.toggle()
    active.value = { row: 5, col: 5 }
    await nextTick()
    expect([v.open.value, v.current.value]).toEqual([false, null])
    expect(pick).not.toHaveBeenCalled()
  })

  it('does not open without a list, and a click on an item enters it', async () => {
    const r = rules(() => ({ prompt: { text: 'только подсказка' } }))
    const { v, pick } = setup(r.source)
    await pass()
    r.calls[0]!.resolve()
    await pass(0)
    expect(v.onKey(key('ArrowDown', { altKey: true }))).toBe(false)
    expect(v.open.value).toBe(false)
    const { v: w, pick: chosen } = await ready()
    expect(w.choose(1)).toBe(true)
    expect(chosen).toHaveBeenCalledWith('Нет', { sheet: 'S', cell: 'A1' })
    expect(pick).not.toHaveBeenCalled()
  })
})

describe('useValidationAlert', () => {
  const alert: SheetValidationAlert = { severity: 'warning', message: 'Лучше больше нуля', sheet: 'S', cell: 'C2' }

  it('shows the alert until the answer, a new alert answering the old with Cancel', async () => {
    const a = useValidationAlert()
    const first = a.ask(alert)
    expect(a.alert.value).toEqual(alert)
    const second = a.ask({ ...alert, severity: 'stop' })
    await expect(first).resolves.toBe('cancel')
    a.answer('retry')
    await expect(second).resolves.toBe('retry')
    expect(a.alert.value).toBeNull()
  })

  it('has Excel’s buttons by style and their answers, the default Retry, No and OK', () => {
    expect(ALERT_BUTTONS.stop.map((b) => `${b.button}:${b.answer}`)).toEqual(['retry:retry', 'cancel:cancel'])
    expect(ALERT_BUTTONS.warning.map((b) => `${b.button}:${b.answer}`)).toEqual(['yes:accept', 'no:retry', 'cancel:cancel'])
    expect(ALERT_BUTTONS.information.map((b) => `${b.button}:${b.answer}`)).toEqual(['ok:accept', 'cancel:cancel'])
    expect(ALERT_DEFAULT).toEqual({ stop: 'retry', warning: 'no', information: 'ok' })
  })
})
