// @vitest-environment node
// localSource over a workbook whose calls outside it the host answers with promises (xlcalc 0.68's `externals.resolve`):
// a cell waiting past its call's `timeoutMs` shows #BUSY! and takes its answer when it arrives, without the user doing anything;
// Refresh All asks again; a formula the engine cannot calculate keeps the value its file stores.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createInlineCalc, type AsyncSession, type BookInput, type Calc, type ExternalCall, type ExternalResolver } from 'xlcalc'
import type { CellUpdate, SourceEvent } from '../src/contract/source'
import { writeXlsx } from 'xlcalc/xlsx'
import { openXlsx, type LocalSource } from '../src/local'
import { closeAll, source, win } from './local-fixture'

let calcs: Calc[] = []
afterEach(async () => {
  await closeAll()
  await Promise.all(calcs.map((c) => c.close()))
  calcs = []
})

/** A1 2, B1 =BAL(A1)*2, C1 =B1+1: BAL is the host's function, answered ten times its argument. */
const balances = (calculation?: BookInput['calculation']): BookInput => ({
  ...(calculation ? { calculation } : {}),
  sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'BAL(A1)*2' }, C1: { formula: 'B1+1' } } }],
})

/** A resolver whose answers wait until the test lets them go (`release`), each call listed. */
function held() {
  const calls: ExternalCall[] = []
  const waiting: (() => void)[] = []
  const resolve: ExternalResolver = (call) => {
    calls.push(call)
    return new Promise((done) => waiting.push(() => done((call.args[0] as number) * 10)))
  }
  const release = () => waiting.splice(0).forEach((go) => go())
  return { calls, resolve, release }
}

/** The values of the source's `changed` events as `sheet!cell=value`, collected from now on. */
function changes(s: LocalSource) {
  const got: CellUpdate[] = []
  const events: SourceEvent[] = []
  s.subscribe((e) => {
    events.push(e)
    if (e.kind === 'changed') got.push(...e.changes)
  })
  const text = () => got.map((c) => `${c.sheet}!${c.cell}=${JSON.stringify(c.value)}`)
  /** Resolves once a `changed` event lists `cell` with its answer, not waiting. */
  const until = (cell: string) =>
    new Promise<void>((done) => {
      const check = () => (got.some((c) => c.cell === cell && !(typeof c.value === 'object' && c.value?.error === 'BUSY')) ? done() : setTimeout(check, 5))
      check()
    })
  return { text, events, until }
}

const BUSY = { error: 'BUSY' }
const row = async (s: LocalSource) => (await win(s, 'S', 1, 1, 1, 3)).values[0]

describe('localSource: answers on their way', () => {
  it('a cell waiting when the workbook opens shows #BUSY! and then its answer', async () => {
    const r = held()
    const s = source(balances(), { session: { timeoutMs: 20, externals: { functions: ['BAL'], resolve: r.resolve } } })
    await s.ready
    expect(await row(s)).toEqual([2, BUSY, BUSY])
    const seen = changes(s)
    r.release()
    await seen.until('B1')
    expect(seen.text()).toEqual(['S!B1=40', 'S!C1=41'])
    expect(await row(s)).toEqual([2, 40, 41])
  })

  it('an edit whose call waits past its deadline shows #BUSY!, and the answer arrives as a change', async () => {
    const r = held()
    const calc = createInlineCalc({ timeoutMs: 20 })
    calcs.push(calc)
    const s = source(balances(), { calc, session: { externals: { functions: ['BAL'], resolve: r.resolve } } })
    await s.ready
    r.release()
    await new Promise((go) => setTimeout(go, 30))
    const seen = changes(s)
    const got = await s.edit([{ sheet: 'S', cell: 'A1', text: '3' }])
    expect(got.ok && got.changes.map((c) => [c.cell, c.value])).toEqual([['A1', 3], ['B1', BUSY], ['C1', BUSY]])
    r.release()
    await seen.until('B1')
    expect(seen.text().slice(-2)).toEqual(['S!B1=60', 'S!C1=61'])
    expect(await row(s)).toEqual([3, 60, 61])
  })

  it('in Manual the answer shows where the cell shows it waiting; the cells reading it wait for a calculation', async () => {
    const r = held()
    const calc = createInlineCalc({ timeoutMs: 20 })
    calcs.push(calc)
    const s = source(balances('manual'), { calc, session: { externals: { functions: ['BAL'], resolve: r.resolve } } })
    await s.ready
    r.release()
    await new Promise((go) => setTimeout(go, 30))
    // B1 asked and shows its answer; C1 read #BUSY! from it and waits for F9
    expect(await row(s)).toEqual([2, 40, BUSY])
    await s.calculate()
    expect(await row(s)).toEqual([2, 40, 41])
    const seen = changes(s)
    // a formula entered calculates and waits, shown as #BUSY!
    const got = await s.edit([{ sheet: 'S', cell: 'D1', text: '=BAL(7)+C1' }])
    expect(got.ok && got.changes.map((c) => [c.cell, c.value])).toEqual([['D1', BUSY]])
    r.release()
    await seen.until('D1')
    expect(seen.text()).toEqual(['S!D1={"error":"BUSY"}', 'S!D1=111'])
    // an edit of A1 leaves B1 as shown; its answer arrives in the session, and B1 still waits for F9
    await s.edit([{ sheet: 'S', cell: 'A1', text: '3' }])
    r.release()
    await new Promise((go) => setTimeout(go, 30))
    expect(await row(s)).toEqual([3, 40, 41])
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: true })
    const now = await s.calculate()
    expect(now.ok && now.changes.map((c) => `${c.cell}=${JSON.stringify(c.value)}`).sort()).toEqual(['B1=60', 'C1=61', 'D1=131'])
  })

  it('closing the source ends the waits: an answer that never comes holds nothing', async () => {
    const resolve: ExternalResolver = () => new Promise(() => {})
    // without a deadline the session waits for the answer while it opens
    const s = source(balances(), { session: { externals: { functions: ['BAL'], resolve } } })
    const seen = changes(s)
    await new Promise((go) => setTimeout(go, 20))
    await s.close()
    expect(seen.events).toEqual([])
  })

  it('closing the source ends a wait that an answer leads to, while opening or during an edit', async () => {
    let asked = 0
    const resolve: ExternalResolver = () => {
      asked++
      return new Promise(() => {})
    }
    // C1 asks once B1 is answered: closing answers B1 with none, and C1 is not asked of the host
    const chained: BookInput = { sheets: [{ name: 'S', cells: { B1: { formula: 'BAL(1)' }, C1: { formula: 'BAL(B1)' } } }] }
    const s = source(chained, { session: { externals: { functions: ['BAL'], resolve } } })
    await new Promise((go) => setTimeout(go, 20))
    const closed = await Promise.race([s.close().then(() => 'closed'), new Promise((go) => setTimeout(() => go('hung'), 1000))])
    expect(closed).toBe('closed')
    expect(asked).toBe(1)
    // an edit waiting for an answer when the source closes ends too
    const r = held()
    const t = source(balances(), { session: { timeoutMs: 20, externals: { functions: ['BAL'], resolve: r.resolve } } })
    await t.ready
    r.release()
    await new Promise((go) => setTimeout(go, 20))
    const edit = t.edit([{ sheet: 'S', cell: 'A1', text: '3' }])
    await new Promise((go) => setTimeout(go, 20))
    const done = await Promise.race([Promise.all([t.close(), edit]).then(() => 'closed'), new Promise((go) => setTimeout(() => go('hung'), 1000))])
    expect(done).toBe('closed')
  })

  it('closing the source while a cell waits sends nothing after it', async () => {
    const r = held()
    const s = source(balances(), { session: { timeoutMs: 20, externals: { functions: ['BAL'], resolve: r.resolve } } })
    await s.ready
    const seen = changes(s)
    await s.close()
    r.release()
    await new Promise((go) => setTimeout(go, 20))
    expect(seen.events).toEqual([])
  })
})

describe('localSource: answers in Manual', () => {
  /** In Manual: A1 2, B1 =BAL(A1)*2, C1 =B1+1, D1 =IFERROR(B1,0); a text argument is answered at once with #N/A. */
  const manualBook = (): BookInput => ({
    calculation: 'manual',
    sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'BAL(A1)*2' }, C1: { formula: 'B1+1' }, D1: { formula: 'IFERROR(B1,0)' } } }],
  })
  const row4 = async (s: LocalSource) => (await win(s, 'S', 1, 1, 1, 4)).values[0]
  function open() {
    const r = held()
    const calc = createInlineCalc({ timeoutMs: 20 })
    calcs.push(calc)
    const resolve: ExternalResolver = (call) => (typeof call.args[0] === 'string' ? { error: 'NA' } : r.resolve(call))
    return { r, s: source(manualBook(), { calc, session: { externals: { functions: ['BAL'], resolve } } }) }
  }

  it('an answer before the first change shows in the cell that asked; the cells reading it wait for F9', async () => {
    const { r, s } = open()
    await s.ready
    expect(await row4(s)).toEqual([2, BUSY, BUSY, 0])
    const seen = changes(s)
    r.release()
    await seen.until('B1')
    expect(seen.text()).toEqual(['S!B1=40'])
    expect(await row4(s)).toEqual([2, 40, BUSY, 0])
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: true })
    await s.calculate()
    expect(await row4(s)).toEqual([2, 40, 41, 40])
  })

  it('a cell showing #BUSY! it read from another waits for F9 after an edit too', async () => {
    const { s } = open()
    await s.ready
    await s.edit([{ sheet: 'S', cell: 'A1', text: 'abc' }])
    expect(await row4(s)).toEqual(['abc', BUSY, BUSY, 0])
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: true })
    await s.calculate()
    expect(await row4(s)).toEqual(['abc', { error: 'NA' }, { error: 'NA' }, 0])
  })
})

describe('localSource: an answer whose calculation runs past the deadline', () => {
  it('is taken again with a longer deadline, rather than leaving its cell waiting', async () => {
    const calc = createInlineCalc({ timeoutMs: 1 })
    calcs.push(calc)
    let go: (v: number[][]) => void = () => {}
    const resolve: ExternalResolver = () => new Promise<number[][]>((done) => (go = done))
    const cells: BookInput['sheets'][number]['cells'] = { A1: { formula: 'BAL(1)' } }
    for (let r = 1; r <= 3000; r++) cells[`Z${r}`] = { formula: `SUM($A$1:$A$3000)+${r}` }
    const s = source({ sheets: [{ name: 'S', cells }] }, { calc, session: { timeoutMs: 5000, externals: { functions: ['BAL'], resolve } } })
    await s.ready
    expect((await win(s, 'S', 1, 1, 1, 1)).values[0]![0]).toEqual(BUSY)
    const seen = changes(s)
    go(Array.from({ length: 3000 }, (_, i) => [i]))
    await seen.until('A1')
    expect((await win(s, 'S', 1, 1, 1, 1)).values[0]![0]).toBe(0)
  }, 30000)
})

describe('localSource: answers that cannot be taken', () => {
  it('are logged, and the sheet reads its windows again', async () => {
    const r = held()
    const inline = createInlineCalc()
    calcs.push(inline)
    // a session whose settleConnections fails as a host's calculator might
    const calc: Calc = {
      compute: (...a) => inline.compute(...a),
      close: async () => {},
      createSession: async (...a) => Object.assign(await inline.createSession(...a), { settleConnections: () => Promise.reject(new Error('lost')) }) as AsyncSession,
    }
    const log = vi.spyOn(console, 'error').mockImplementation(() => {})
    const s = source(balances(), { calc, session: { timeoutMs: 20, externals: { functions: ['BAL'], resolve: r.resolve } } })
    await s.ready
    const seen = changes(s)
    r.release()
    await new Promise((go) => setTimeout(go, 30))
    expect(log).toHaveBeenCalledTimes(1)
    expect(String(log.mock.calls[0]![1])).toContain('lost')
    expect(seen.events.map((e) => e.kind)).toEqual(['reload'])
    log.mockRestore()
  })
})

describe('localSource: Refresh All', () => {
  /** A resolver answering at once, a new answer each time it is asked: the call's argument times ten plus the asks so far. */
  function counting() {
    let asks = 0
    const resolve: ExternalResolver = ({ args }) => (args[0] as number) * 10 + asks++
    return { resolve, asks: () => asks }
  }

  it('asks every call again and sends what changed, not as a step of undo', async () => {
    const r = counting()
    const s = source(balances(), { session: { externals: { functions: ['BAL'], resolve: r.resolve } } })
    await s.ready
    expect(await row(s)).toEqual([2, 40, 41])
    await s.edit([{ sheet: 'S', cell: 'D1', text: 'x' }])
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    const seen = changes(s)
    const got = await s.refreshConnections!()
    expect(got.ok && got.changes.map((c) => `${c.cell}=${JSON.stringify(c.value)}`)).toEqual(['B1=42', 'C1=43'])
    expect(seen.text()).toEqual(['S!B1=42', 'S!C1=43'])
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    // the step of undo is still the edit's
    await s.undo()
    expect((await win(s, 'S', 1, 1, 1, 4)).values[0]).toEqual([2, 42, 43, null])
  })

  it('asks again only the functions named', async () => {
    const r = counting()
    const book = balances()
    book.sheets[0]!.cells.E1 = { formula: 'OTHER(1)' }
    const s = source(book, { session: { externals: { functions: ['BAL', 'OTHER'], resolve: r.resolve } } })
    await s.ready
    const got = await s.refreshConnections!(['OTHER'])
    expect(got.ok && got.changes.map((c) => `${c.cell}=${JSON.stringify(c.value)}`)).toEqual(['E1=12'])
  })

  it('answers on their way arrive after it as changes', async () => {
    const r = held()
    const calc = createInlineCalc({ timeoutMs: 20 })
    calcs.push(calc)
    const s = source(balances(), { calc, session: { externals: { functions: ['BAL'], resolve: r.resolve } } })
    await s.ready
    r.release()
    await new Promise((go) => setTimeout(go, 30))
    const seen = changes(s)
    const got = await s.refreshConnections!()
    expect(got.ok && got.changes.map((c) => [c.cell, c.value])).toEqual([['B1', BUSY], ['C1', BUSY]])
    expect(r.calls.length).toBe(2)
    r.release()
    await seen.until('B1')
    expect(await row(s)).toEqual([2, 40, 41])
  })

  it('in Manual the cells it calculated wait for a calculation', async () => {
    const r = counting()
    const s = source(balances('manual'), { session: { externals: { functions: ['BAL'], resolve: r.resolve } } })
    await s.ready
    const got = await s.refreshConnections!()
    expect(got.ok && got.changes).toEqual([])
    expect(await row(s)).toEqual([2, 40, 41])
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: true })
    await s.calculate()
    expect(await row(s)).toEqual([2, 42, 43])
  })

  it('asks the calls again only with a resolver: without one it asks nothing, the PivotTables being all it refreshes', async () => {
    // values given in advance answer nothing anew: asking would only drop a cube's value the file stores
    const s = source(balances(), { session: { externals: { functions: ['BAL'], values: { 'BAL(2)': 5 } } } })
    const before = await row(s)
    expect(await s.refreshConnections()).toEqual({ ok: true, changes: [] })
    expect(await row(s)).toEqual(before)
  })
})

describe('localSource: the values a file stores', () => {
  /** A1 2, B1 =MYADDIN(A1)*2 (an add-in's function), C1 =B1+1, D1 a cube's value; the file stores 40, 41 and 7. */
  const addIn = (): BookInput => ({
    sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'MYADDIN(A1)*2' }, C1: { formula: 'B1+1' }, D1: { formula: 'CUBEVALUE("Sales","[Measures].[Amount]")' } } }],
  })
  const cached = { sheets: [{ name: 'S', cells: { B1: { value: 40 }, C1: { value: 41 }, D1: { value: 7 } } }] }
  const row4 = async (s: LocalSource) => (await win(s, 'S', 1, 1, 1, 4)).values[0]

  it("a formula over a function the engine does not know keeps the file's value until it is calculated, as Excel shows it", async () => {
    const s = source(addIn(), { session: { cached } })
    expect(await row4(s)).toEqual([2, 40, 41, 7])
    // an edit of what it reads calculates it, as Excel without the add-in
    await s.edit([{ sheet: 'S', cell: 'A1', text: '3' }])
    expect(await row4(s)).toEqual([3, { error: 'NAME' }, { error: 'NAME' }, 7])
  })

  it("the host's unknownFunctions wins", async () => {
    const s = source(addIn(), { session: { cached, unknownFunctions: 'name' } })
    expect(await row4(s)).toEqual([2, { error: 'NAME' }, { error: 'NAME' }, 7])
  })

  it("a file read by openXlsx keeps its values, and CELL and INFO know the file's name", async () => {
    const book = addIn()
    book.sheets[0]!.cells.E1 = { formula: 'CELL("filename",A1)' }
    const { data } = await writeXlsx({ book, result: { sheets: [{ name: 'S', cells: { ...cached.sheets[0]!.cells, E1: { value: '' } } }] } })
    const read = await openXlsx(new File([data as BlobPart], 'Budget.xlsx'))
    expect(read.file).toEqual({ name: 'Budget.xlsx' })
    const s = source(read.book, { session: { cached: read.cached, file: read.file } })
    expect((await win(s, 'S', 1, 1, 1, 5)).values[0]).toEqual([2, 40, 41, 7, '[Budget.xlsx]S'])
    // bytes have no name
    expect((await openXlsx(data)).file).toBeUndefined()
  })
})
