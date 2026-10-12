// @vitest-environment node
// localSource and Excel's outline: the sheets give their groups from xlcalc's layout, and each command is one setLayout
// call (one step of the undo) answered by a reload.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { LocalSource } from '../src/local'
import type { OutlineChange, SourceEvent } from '../src/source'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

/** A command without its sheet, each kind of `OutlineChange` kept apart. */
type Change = OutlineChange extends infer C ? (C extends unknown ? Omit<C, 'sheet'> : never) : never

/** Rows 2:10 in two levels (2:4 and 6:9 inside), 6:9 collapsed under row 10; columns B:D grouped; a SUBTOTAL over the rows. */
const book = (extra: Partial<BookInput['sheets'][number]> = {}): BookInput => ({
  locale: 'en-US',
  sheets: [
    {
      name: 'S',
      cells: Object.fromEntries([...Array.from({ length: 11 }, (_, i) => [`A${i + 1}`, { value: i + 1 }]), ['C12', { formula: 'SUBTOTAL(109,A1:A11)' }]]),
      rows: { '2:4': { outlineLevel: 2 }, '5': { outlineLevel: 1 }, '6:9': { outlineLevel: 2, hidden: true }, '10': { outlineLevel: 1, collapsed: true } },
      columns: { 'B:D': { outlineLevel: 1 } },
      ...extra,
    },
    { name: 'T', cells: { A1: { value: 1 } } },
  ],
})
async function opened(b: BookInput = book()) {
  const s = source(b)
  const events: SourceEvent[] = []
  s.subscribe((e) => events.push(e))
  return { s, events, run: (change: Change) => s.outline({ sheet: 'S', ...change } as OutlineChange) }
}
const outlineOf = async (s: LocalSource) => (await s.sheets())[0]!.outline
const hiddenRows = async (s: LocalSource) => Object.entries((await win(s, 'S', 1, 12, 1, 1)).rowHeights).filter(([, h]) => h === 0).map(([r]) => Number(r))
const subtotal = async (s: LocalSource) => (await win(s, 'S', 12, 1, 3, 1)).values[0]![0]

describe('localSource: the outline of a sheet', () => {
  it('gives the groups from the levels and hidden lines, Excel’s summaries by default; a sheet without groups has none', async () => {
    const { s } = await opened()
    expect(await outlineOf(s)).toEqual({
      rows: [
        { from: 2, to: 10, level: 1, collapsed: false },
        { from: 2, to: 4, level: 2, collapsed: false },
        { from: 6, to: 9, level: 2, collapsed: true },
      ],
      columns: [{ from: 2, to: 4, level: 1, collapsed: false }],
      summaryBelow: true,
      summaryRight: true,
    })
    expect((await s.sheets())[1]).not.toHaveProperty('outline')
  })

  it('takes the summary settings and a height of 0 as hidden', async () => {
    const { s } = await opened(book({ rows: { '3:4': { outlineLevel: 1, height: 0 } }, columns: {}, outline: { summaryBelow: false } }))
    expect(await outlineOf(s)).toEqual({ rows: [{ from: 3, to: 4, level: 1, collapsed: true }], columns: [], summaryBelow: false, summaryRight: true })
  })
})

describe('localSource: outline commands', () => {
  it('Show Detail of the outer group keeps the inner one collapsed; each command is one step of the undo, sent as a reload', async () => {
    const { s, events, run } = await opened()
    expect(await hiddenRows(s)).toEqual([6, 7, 8, 9])
    expect(await run({ op: 'collapse', rows: '2:10', level: 1 })).toEqual({ ok: true, changes: [expect.objectContaining({ cell: 'C12' })] })
    expect(events).toEqual([{ kind: 'reload' }])
    expect(await hiddenRows(s)).toEqual([2, 3, 4, 5, 6, 7, 8, 9, 10])
    expect(await subtotal(s)).toBe(1 + 11)
    expect((await s.book()).sheets[0]!.rows!['11']).toEqual({ collapsed: true })
    await run({ op: 'expand', rows: '2:10', level: 1 })
    expect(await hiddenRows(s)).toEqual([6, 7, 8, 9])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
    await s.undo()
    expect(await hiddenRows(s)).toEqual([2, 3, 4, 5, 6, 7, 8, 9, 10])
  })

  it('Group and Ungroup over columns; a level button over rows; a refusal is the rule’s message', async () => {
    const { s, run } = await opened()
    expect(await run({ op: 'group', columns: 'C:C' })).toMatchObject({ ok: true })
    expect((await outlineOf(s))!.columns).toEqual([
      { from: 2, to: 4, level: 1, collapsed: false },
      { from: 3, to: 3, level: 2, collapsed: false },
    ])
    await run({ op: 'ungroup', columns: 'B:D' })
    expect((await outlineOf(s))!.columns).toEqual([{ from: 3, to: 3, level: 1, collapsed: false }])
    expect(await run({ op: 'ungroup', columns: 'F:G' })).toEqual({ ok: false, reason: 'refused', message: 'The columns F:G are in no group' })
    await run({ op: 'showLevel', axis: 'rows', level: 3 })
    expect(await hiddenRows(s)).toEqual([])
    expect(await run({ op: 'collapse', rows: '3:4', level: 2 })).toEqual({ ok: false, reason: 'refused', message: 'No group of level 2 over the rows 3:4' })
  })

  it('a line hidden by a height of 0 is shown again at the default height', async () => {
    const { s, run } = await opened(book({ rows: { '3:4': { outlineLevel: 1, height: 0 } }, columns: {} }))
    await run({ op: 'expand', rows: '3:4', level: 1 })
    expect(await hiddenRows(s)).toEqual([])
    expect((await s.book()).sheets[0]!.rows ?? {}).toEqual({ '3:4': { outlineLevel: 1 } })
  })

  it('an unknown sheet fails', async () => {
    const { s } = await opened()
    expect(await s.outline({ op: 'showLevel', sheet: 'Nope', axis: 'rows', level: 1 })).toMatchObject({ ok: false, reason: 'failed' })
  })
})
