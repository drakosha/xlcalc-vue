// Excel's commands in the cell menu and as the sheet's methods: Reapply for the display filter (Ctrl+Alt+L), Freeze Panes
// through the source's `freeze` with the lock as a view setting over it, and Trace Precedents as a panel next to the grid.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createSession, registerLocale, type BookInput } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { memorySource, ruMessages, type MemoryBook, type SheetViewApi } from '../src'
import { EXPLAIN_DELAY_MS, explainLines } from '../src/core/explain'
import { localSource } from '../src/local'
import type { EditOutcome } from '../src/contract/source'
import { active, address, goTo, key, mountSheet, settle, until } from './helpers'

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
  document.body.innerHTML = ''
  localStorage.clear()
})

const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const items = () => [...document.querySelectorAll<HTMLElement>('[data-xs-menu]')].map((b) => b.dataset.xsMenu)
const label = (id: string) => document.querySelector(`[data-xs-menu="${id}"]`)?.textContent?.trim()
async function context(w: VueWrapper, cell: string): Promise<void> {
  await w.find(`[data-cell="${cell}"]`).trigger('contextmenu', { button: 2, clientX: 200, clientY: 100 })
  await settle()
}
async function pick(id: string): Promise<void> {
  document.querySelector<HTMLElement>(`[data-xs-menu="${id}"]`)!.click()
  await settle()
}
const numbers = (w: VueWrapper) => w.findAll('.xs-num').map((n) => n.text())
const td = (w: VueWrapper, row: number, col: number) => w.find(`[data-cell="${row}:${col}"]`)
/** Fake timers for the panel's pause; setImmediate stays real, since flushPromises needs it. */
const fakeTimers = () => vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'Date'] })

describe('Reapply for the display filter', () => {
  /** A list filtered on B = a under the host's storage key: rows 2, 4 and 6 visible. */
  const book = (): MemoryBook => ({
    sheets: [{ name: 'S', cells: { A1: 'Name', B1: 'Status', A2: 'n2', B2: 'a', A3: 'n3', B3: 'b', A4: 'n4', B4: 'a', A5: 'n5', B5: 'b', A6: 'n6', B6: 'a' } }],
  })
  async function filtered() {
    localStorage.setItem('k:filter:S', JSON.stringify({ cols: { 2: { values: ['a'] } }, sort: null }))
    const source = memorySource(book())
    const visibleRows = vi.spyOn(source, 'visibleRows')
    const w = await mountSheet(source, { storageKey: 'k' })
    expect(numbers(w)).toEqual(['1', '2', '4', '6'])
    return { source, visibleRows, w }
  }

  it('Ctrl+Alt+L reapplies the conditions: a row that no longer matches after an edit is hidden, one that now matches appears; the scroll and a visible active cell stay', async () => {
    const { source, visibleRows, w } = await filtered()
    await td(w, 6, 1).trigger('pointerdown', { button: 0 })
    source.setCells('S', { B4: 'b', B3: 'a' })
    await settle()
    // Until Reapply the rows stay, as Excel keeps them.
    expect(numbers(w)).toEqual(['1', '2', '4', '6'])
    const scroller = w.find('.xs-scroller').element as HTMLElement
    scroller.scrollTop = 12
    const calls = visibleRows.mock.calls.length
    await key(w, 'l', { ctrlKey: true, altKey: true, code: 'KeyL' })
    await settle()
    expect(visibleRows.mock.calls.length).toBe(calls + 1)
    expect(numbers(w)).toEqual(['1', '2', '3', '6'])
    expect(active(w)).toBe('6:1')
    expect(scroller.scrollTop).toBe(12)
    // On a Mac Option turns the key into another character: the key's position counts.
    source.setCells('S', { B2: 'b' })
    await key(w, '¬', { ctrlKey: true, altKey: true, code: 'KeyL' })
    await settle()
    expect(numbers(w)).toEqual(['1', '3', '6'])
  })

  it('an active cell whose row is hidden moves to the row now at its position; the selection collapses to it', async () => {
    const { source, w } = await filtered()
    await goTo(w, 'B4:C6')
    source.setCells('S', { B4: 'b' })
    await settle()
    await key(w, 'l', { ctrlKey: true, altKey: true, code: 'KeyL' })
    await settle()
    expect(numbers(w)).toEqual(['1', '2', '6'])
    expect([active(w), address(w)]).toEqual(['6:2', 'B6'])
  })

  it("the cell menu, the filter status line's button and the host's reapplyFilter() reapply too; with no filter there is nothing to reapply", async () => {
    const plain = memorySource(book())
    const plainRows = vi.spyOn(plain, 'visibleRows')
    const bare = await mountSheet(plain)
    await context(bare, '2:1')
    expect(items()).not.toContain('reapply')
    expect(await api(bare).reapplyFilter()).toBe(false)
    await key(bare, 'l', { ctrlKey: true, altKey: true, code: 'KeyL' })
    expect(plainRows).not.toHaveBeenCalled()
    expect(bare.find('[data-xs-filter-reapply]').exists()).toBe(false)
    document.body.innerHTML = ''

    const { source, visibleRows, w } = await filtered()
    await context(w, '2:1')
    // no Freeze Panes under the filter: its rows on screen are not the sheet's
    expect(items()).toEqual(['cut', 'copy', 'paste', 'reapply', 'precedents', 'defineName'])
    source.setCells('S', { B2: 'b' })
    await pick('reapply')
    expect(numbers(w)).toEqual(['1', '4', '6'])
    source.setCells('S', { B3: 'a' })
    await w.find('[data-xs-filter-reapply]').trigger('click')
    await settle()
    expect(numbers(w)).toEqual(['1', '3', '4', '6'])
    source.setCells('S', { B6: 'b' })
    const calls = visibleRows.mock.calls.length
    expect(await api(w).reapplyFilter()).toBe(true)
    await settle()
    expect(visibleRows.mock.calls.length).toBe(calls + 1)
    expect(numbers(w)).toEqual(['1', '3', '4'])
  })

  it("uses Excel's names in Russian", async () => {
    localStorage.setItem('k:filter:S', JSON.stringify({ cols: { 2: { values: ['a'] } }, sort: null }))
    const w = await mountSheet(memorySource(book()), { storageKey: 'k', messages: ruMessages })
    // the stored filter is reapplied through the source's visibleRows, in as many rounds as the machine's load takes
    await until(() => w.find('[data-xs-filter-reapply]').exists())
    await context(w, '2:1')
    expect([label('reapply'), label('precedents')]).toEqual(['Применить повторно', 'Влияющие ячейки'])
    expect(w.find('[data-xs-filter-reapply]').attributes('title')).toBe(ruMessages.filterReapply)
    document.body.innerHTML = ''
    const plain = await mountSheet(memorySource(book()), { messages: ruMessages })
    await context(plain, '2:1')
    expect([label('freeze'), label('freezeTopRow'), label('freezeFirstColumn')]).toEqual(['Закрепить области', 'Закрепить верхнюю строку', 'Закрепить первый столбец'])
  })
})

describe('Freeze Panes', () => {
  const book = (freeze?: { rows: number; cols: number }): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'h', H90: 'end' }, ...(freeze ? { freeze } : {}) }, { name: 'T', cells: { A1: 't' } }] })
  const frozen = (w: VueWrapper) => ({ corner: td(w, 1, 1).classes().includes('xs-fz-corner'), rows: w.findAll('.xs-fz-row').length > 0, cols: w.findAll('.xs-fz-col').length > 0 })

  it('the menu freezes the rows above and the columns to the left of the active cell through `freeze`; Unfreeze Panes removes them', async () => {
    const source = memorySource(book())
    const freeze = vi.spyOn(source, 'freeze')
    const sheets = vi.spyOn(source, 'sheets')
    const w = await mountSheet(source)
    expect(frozen(w)).toEqual({ corner: false, rows: false, cols: false })
    await context(w, '4:3')
    expect(label('freeze')).toBe('Freeze Panes')
    const reads = sheets.mock.calls.length
    await pick('freeze')
    expect(freeze).toHaveBeenLastCalledWith('S', { rows: 3, cols: 2 })
    // The sheet is read again, as after a structure operation.
    expect(sheets.mock.calls.length).toBe(reads + 1)
    expect(frozen(w)).toEqual({ corner: true, rows: true, cols: true })
    expect(td(w, 3, 2).classes()).toContain('xs-fz-corner')
    expect(active(w)).toBe('4:3')
    await context(w, '10:5')
    expect(items()).toContain('unfreeze')
    expect(label('unfreeze')).toBe('Unfreeze Panes')
    await pick('unfreeze')
    expect(freeze).toHaveBeenLastCalledWith('S', null)
    expect(frozen(w)).toEqual({ corner: false, rows: false, cols: false })
    await context(w, '10:5')
    await pick('freezeTopRow')
    expect(freeze).toHaveBeenLastCalledWith('S', { rows: 1, cols: 0 })
    expect(frozen(w)).toEqual({ corner: false, rows: true, cols: false })
    // Excel's Freeze First Column replaces the previous panes.
    await context(w, '10:5')
    await pick('freezeFirstColumn')
    expect(freeze).toHaveBeenLastCalledWith('S', { rows: 0, cols: 1 })
    expect(frozen(w)).toEqual({ corner: false, rows: false, cols: true })
  })

  it("panes frozen by a command display regardless of the lock that hid the workbook's panes and of the half-viewport rule; unfreezing clears the lock's choice", async () => {
    const source = memorySource(book({ rows: 2, cols: 1 }))
    const w = await mountSheet(source, { storageKey: 'k' })
    await w.find('.xs-tools .xs-tool').trigger('click')
    expect(localStorage.getItem('k:freeze:S')).toBe('false')
    expect(frozen(w).corner).toBe(false)
    await goTo(w, 'C4')
    expect(await api(w).freezePanes()).toEqual({ ok: true, changes: [] })
    await settle()
    expect(frozen(w)).toEqual({ corner: true, rows: true, cols: true })
    expect(localStorage.getItem('k:freeze:S')).toBe('true')
    // 40 rows take over half of the 600 px viewport: frozen by the command, they are displayed all the same.
    await goTo(w, 'A41')
    await api(w).freezePanes('panes')
    await settle()
    expect(frozen(w)).toEqual({ corner: false, rows: true, cols: false })
    await api(w).freezePanes('unfreeze')
    await settle()
    expect(localStorage.getItem('k:freeze:S')).toBeNull()
    expect(w.findAll('.xs-fz-row')).toHaveLength(0)
  })

  it("with A1 active it freezes at the viewport's middle, as Excel does; entire columns or rows freeze to the left of or above them", async () => {
    const source = memorySource(book())
    const freeze = vi.spyOn(source, 'freeze')
    const w = await mountSheet(source)
    // The viewport is 800 × 600 in happy-dom: the letters take 24 px and the numbers 56 px; rows are 20 px, columns 71 px.
    await api(w).freezePanes()
    expect(freeze).toHaveBeenLastCalledWith('S', { rows: 14, cols: 5 })
    await w.findAll('.xs-head')[2]!.trigger('click')
    await api(w).freezePanes()
    expect(freeze).toHaveBeenLastCalledWith('S', { rows: 0, cols: 2 })
    await w.findAll('.xs-num').find((n) => n.text() === '5')!.trigger('click')
    await api(w).freezePanes()
    expect(freeze).toHaveBeenLastCalledWith('S', { rows: 4, cols: 0 })
  })

  it('nothing is sent without `freeze`, on a read-only sheet, or with nothing to unfreeze; a rejection is displayed', async () => {
    const source = memorySource(book())
    const freeze = vi.spyOn(source, 'freeze')
    const ro = await mountSheet(source, { readOnly: true })
    await context(ro, '3:3')
    expect(items()).toEqual(['copy', 'precedents'])
    expect(await api(ro).freezePanes()).toBeNull()
    document.body.innerHTML = ''
    const w = await mountSheet(source)
    expect(await api(w).freezePanes('unfreeze')).toBeNull()
    expect(freeze).not.toHaveBeenCalled()
    freeze.mockResolvedValueOnce({ ok: false, reason: 'refused', message: 'Нельзя закрепить' })
    await goTo(w, 'B2')
    await api(w).freezePanes()
    await settle()
    expect(w.find('.xs-message').text()).toBe('Нельзя закрепить')
    document.body.innerHTML = ''
    const none = memorySource(book())
    delete (none as { freeze?: unknown }).freeze
    const bare = await mountSheet(none)
    await context(bare, '3:3')
    expect(items()).toEqual(['cut', 'copy', 'paste', 'precedents', 'defineName'])
    expect(await api(bare).freezePanes('topRow')).toBeNull()
  })

  it('with the local source the workbook stores the panes as a single undo step: Ctrl+Z removes them', async () => {
    const source = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells: { A1: { value: 1 }, C5: { value: 2 } } }] }, { calc: 'inline' })
    const w = await mountSheet(source)
    await settle(12)
    await context(w, '3:2')
    await pick('freeze')
    await settle(12)
    expect((await source.sheets())[0]!.freeze).toEqual({ rows: 2, cols: 1 })
    expect(frozen(w)).toEqual({ corner: true, rows: true, cols: true })
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    expect((await source.sheets())[0]!.freeze).toBeUndefined()
    expect(frozen(w)).toEqual({ corner: false, rows: false, cols: false })
    await source.close()
  })
})

describe('the Trace Precedents panel', () => {
  const book = (): MemoryBook => ({
    sheets: [
      { name: 'Data', cells: { A2: 10, A3: 20, A4: { value: 30, formula: 'SUM(A2:A3)' }, B2: { value: 60, formula: 'A4*2' } } },
      { name: 'Q 4', cells: { A1: { value: 61, formula: 'Data!B2+1' }, A2: { value: 122, formula: 'A1*2' } } },
    ],
  })
  /** An `explain` whose responses the test provides: each call waits for its `answer`. */
  function held() {
    const calls: { sheet: string; cell: string; answer: (text: string) => void }[] = []
    const explain = vi.fn((sheet: string, cell: string) => new Promise<string>((answer) => calls.push({ sheet, cell, answer })))
    return { calls, explain }
  }
  const panel = (w: VueWrapper) => w.find('[data-xs-explain]')
  /** The panel's pause after the active cell moved, and the response. */
  async function follow(): Promise<void> {
    await new Promise((r) => setTimeout(r, EXPLAIN_DELAY_MS + 30))
    await settle()
  }
  const lines = (w: VueWrapper) => w.findAll('[data-xs-explain-line]').map((l) => l.text())

  it('Trace Precedents opens the panel for the active cell; it follows the cell after a pause and discards a late response', async () => {
    fakeTimers()
    const source = memorySource(book())
    const { calls, explain } = held()
    source.explain = explain
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    await context(w, '2:2')
    expect(label('precedents')).toBe('Trace Precedents')
    await pick('precedents')
    await vi.advanceTimersByTimeAsync(0)
    expect(panel(w).exists()).toBe(true)
    expect(calls.map((c) => c.cell)).toEqual(['B2'])
    expect(panel(w).text()).toContain('Loading…')
    // Held-down keys request once, after the pause.
    await key(w, 'ArrowDown')
    await vi.advanceTimersByTimeAsync(EXPLAIN_DELAY_MS - 50)
    await key(w, 'ArrowDown')
    await vi.advanceTimersByTimeAsync(EXPLAIN_DELAY_MS - 50)
    expect(calls).toHaveLength(1)
    await vi.advanceTimersByTimeAsync(60)
    expect(calls.map((c) => c.cell)).toEqual(['B2', 'B4'])
    calls[1]!.answer('Data!B4 = empty')
    await settle()
    calls[0]!.answer('Data!B2 = 60  ← A4*2')
    await settle()
    expect(w.find('[data-xs-explain-place]').text()).toBe('Data!B4')
    expect(lines(w)).toEqual(['Data!B4 = empty'])
    // A value changed by an edit requests again.
    source.setCells('Data', { B4: '5' })
    await vi.advanceTimersByTimeAsync(EXPLAIN_DELAY_MS + 10)
    expect(calls.map((c) => c.cell)).toEqual(['B2', 'B4', 'B4'])
    calls[2]!.answer('Data!B4 = 5')
    await settle()
    expect(lines(w)).toEqual(['Data!B4 = 5'])
  })

  it("shows the in-memory workbook's tree; a reference goes to its range, one without a sheet name on its line's sheet; Esc, the close button and the host close it", async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, "'Q 4'!A2")
    expect(api(w).explain()).toBe(true)
    await settle()
    expect(w.find('[data-xs-explain-place]').text()).toBe("'Q 4'!A2")
    expect(lines(w)).toEqual(["'Q 4'!A2 = 122  ← A1*2", "'Q 4'!A1 = 61  ← Data!B2+1", 'Data!B2 = 60  ← A4*2', 'Data!A4 = 30  ← SUM(A2:A3)'])
    expect(w.findAll('[data-xs-explain-line]').map((l) => l.attributes('style'))).toEqual([0, 1, 2, 3].map((d) => `--xs-explain-depth: ${d};`))
    // A4 on Data's line is Data's A4.
    const a4 = w.findAll('[data-xs-explain-line]')[2]!.find('[data-xs-explain-ref="A4"]')
    expect(a4.attributes('title')).toBe('Go to sheet “Data”')
    await a4.trigger('click')
    await follow()
    expect([w.find('.xs-tab-on').text(), active(w)]).toEqual(['Data', '4:1'])
    expect(w.find('[data-xs-explain-place]').text()).toBe('Data!A4')
    expect(lines(w)).toEqual(['Data!A4 = 30  ← SUM(A2:A3)', 'Data!A2:A3 (2 cells)'])
    await w.find('[data-xs-explain-ref="Data!A2:A3"]').trigger('click')
    await follow()
    expect(address(w)).toBe('A2:A3')
    await key(w, 'Escape')
    expect(panel(w).exists()).toBe(false)
    expect(api(w).explain()).toBe(true)
    await settle()
    await w.find('[data-xs-explain-close]').trigger('click')
    expect(panel(w).exists()).toBe(false)
    expect(document.activeElement).toBe(w.find('.xs-scroller').element)
    api(w).explain()
    await settle()
    await w.find('[data-xs-explain-ref="Data!A2"]').trigger('keydown', { key: 'Escape' })
    expect(panel(w).exists()).toBe(false)
  })

  it('a failed response is reported; without `explain` there is no menu item and no panel', async () => {
    const source = memorySource(book())
    source.explain = vi.fn(async () => Promise.reject(new Error('boom')))
    const w = await mountSheet(source, { messages: ruMessages })
    api(w).explain()
    await settle()
    expect(panel(w).find('[role=alert]').text()).toBe('Не удалось получить влияющие ячейки')
    document.body.innerHTML = ''
    const none = memorySource(book())
    delete (none as { explain?: unknown }).explain
    const bare = await mountSheet(none)
    await context(bare, '2:2')
    expect(items()).not.toContain('precedents')
    expect(api(bare).explain()).toBe(false)
    expect(panel(bare).exists()).toBe(false)
  })
})

describe('the explain text', () => {
  const input = (locale: string): BookInput => ({
    locale,
    sheets: [
      { name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { formula: 'SUM(A1:A2)' }, B1: { formula: 'A3*2' } } },
      { name: 'T 2', cells: { A1: { formula: 'S!B1+1' }, A2: { formula: 'ROUND(A1*1.5,0)' } } },
    ],
  })
  const shape = (text: string, sheet: string) => explainLines(text, sheet).map((l) => [l.depth, l.parts.map((p) => (p.ref ? `[${p.ref.sheet}|${p.text}]` : p.from ? `{${p.text}}` : p.text)).join('')])

  it("parses xlcalc's text: the depth by the indent, each reference with its sheet, the formula after the arrow separated", () => {
    registerLocale(ruRU)
    const text = createSession(input('en-US')).explain('T 2', 'A1').text
    expect(shape(text, 'T 2')).toEqual([
      [0, "[T 2|'T 2'!A1] = 7{  ← }[S|S!B1]{+1}"],
      [1, '[S|S!B1] = 6{  ← }[S|A3]{*2}'],
      [2, '[S|S!A3] = 3{  ← SUM(}[S|A1:A2]{)}'],
      [3, '[S|S!A1:A2] (2 cells)'],
    ])
    // ru-RU's words, numbers and separators parse the same; a reference without a sheet name is on its line's sheet.
    const ru = createSession(input('ru-RU')).explain('T 2', 'A2').text
    expect(ru.split('\n')[0]).toBe("'T 2'!A2 = 11  ← ОКРУГЛ(A1*1,5;0)")
    expect(shape(ru, 'S')[0]).toEqual([0, "[T 2|'T 2'!A2] = 11{  ← ОКРУГЛ(}[T 2|A1]{*1,5;0)}"])
  })

  it("memorySource writes a tree like xlcalc's en-US one, from the formulas as given", async () => {
    const memory = memorySource({
      sheets: [
        { name: 'S', cells: { A1: 1, A2: 2, A3: { value: 3, formula: 'SUM(A1:A2)' }, B1: { value: 6, formula: 'A3*2' }, C1: 'x', D1: { value: 9, formula: 'SUM(A:A)+C1' } } },
        { name: 'T 2', cells: { A1: { value: 7, formula: 'S!B1+1' } } },
      ],
    })
    const session = createSession(input('en-US'))
    expect(await memory.explain!('T 2', 'A1')).toBe(session.explain('T 2', 'A1').text)
    expect(await memory.explain!('S', 'C1')).toBe('S!C1 = "x"')
    expect(await memory.explain!('S', 'E5')).toBe(session.explain('S', 'E5').text)
    expect(await memory.explain!('S', 'D1')).toBe('S!D1 = 9  ← SUM(A:A)+C1\n  S!A:A (3 cells)\n  S!C1 = "x"')
  })
})

describe('memorySource under a display filter', () => {
  it('the grid keeps its rows when a window is re-read after an edit made one fail the conditions', async () => {
    localStorage.setItem('k:filter:S', JSON.stringify({ cols: { 2: { values: ['a'] } }, sort: null }))
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 'Name', B1: 'Status', A2: 'n2', B2: 'a', A3: 'n3', B3: 'b', A4: 'n4', B4: 'a', A5: 'n5', B5: 'b', A6: 'n6', B6: 'a' } }] })
    const w = await mountSheet(source, { storageKey: 'k' })
    source.setCells('S', { B4: 'b' })
    source.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1:B6' }] })
    await settle()
    expect(numbers(w)).toEqual(['1', '2', '4', '6'])
    expect([2, 4, 6].map((r) => td(w, r, 1).text())).toEqual(['n2', 'n4', 'n6'])
  })
})

describe('memorySource’s freeze', () => {
  it('keeps the panes in its own copy of the sheet', async () => {
    const book: MemoryBook = { sheets: [{ name: 'S', cells: {}, freeze: { rows: 1, cols: 0 } }] }
    const source = memorySource(book)
    const outcome: EditOutcome = await source.freeze!('S', { rows: 2, cols: 3 })
    expect(outcome).toEqual({ ok: true, changes: [] })
    expect((await source.sheets())[0]!.freeze).toEqual({ rows: 2, cols: 3 })
    await source.freeze!('S', { rows: 0, cols: 0 })
    expect((await source.sheets())[0]!.freeze).toBeUndefined()
    expect(book.sheets[0]!.freeze).toEqual({ rows: 1, cols: 0 })
    expect(await source.freeze!('X', null)).toMatchObject({ ok: false, reason: 'failed' })
  })
})
