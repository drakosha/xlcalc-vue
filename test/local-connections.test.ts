// @vitest-environment node
// localSource over a workbook whose calls outside it the host answers with promises (xlcalc 0.68's `externals.resolve`):
// a cell waiting past its call's `timeoutMs` shows #BUSY! and takes its answer when it arrives, without the user doing anything.
import { afterEach, describe, expect, it } from 'vitest'
import { createInlineCalc, type BookInput, type Calc, type ExternalCall, type ExternalResolver } from 'xlcalc'
import type { CellUpdate, SourceEvent } from '../src/contract/source'
import type { LocalSource } from '../src/local'
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
