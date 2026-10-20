// @vitest-environment node
// localSource and Excel's outline on xlcalc's groups and commands (its docs/outline.md): the sheets return `outlineGroups`,
// each command is one `editStructure` (a single undo step, SUBTOTAL recalculated) followed by a reload.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import { openXlsx, saveXlsx, type LocalSource } from '../src/local'
import type { OutlineChange, SourceEvent } from '../src/contract/source'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

/** A command without its sheet, with each kind of `OutlineChange` kept separate. */
type Change = OutlineChange extends infer C ? (C extends unknown ? Omit<C, 'sheet'> : never) : never
type Sheet = BookInput['sheets'][number]

/** Rows 2:10 in two levels (2:4 and 6:9 inside), 6:9 collapsed under row 10; columns B:D grouped; a SUBTOTAL over the rows. */
const book = (extra: Partial<Sheet> = {}, locale = 'en-US'): BookInput => ({
  locale,
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
/** Rows 3:10 at level 1 with 5:7 at level 2 inside: groups 3:10 (summary row 11) and 5:7 (summary row 8). */
const NESTED: Sheet['rows'] = { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 }, '8:10': { outlineLevel: 1 } }
async function opened(b: BookInput = book()) {
  const s = source(b)
  const events: SourceEvent[] = []
  s.subscribe((e) => events.push(e))
  return { s, events, run: (change: Change) => s.outline({ sheet: 'S', ...change } as OutlineChange) }
}
const outlineOf = async (s: LocalSource) => (await s.sheets())[0]!.outline
const hiddenRows = async (s: LocalSource) => Object.entries((await win(s, 'S', 1, 12, 1, 1)).rowHeights).filter(([, h]) => h === 0).map(([r]) => Number(r))
const subtotal = async (s: LocalSource) => (await win(s, 'S', 12, 1, 3, 1)).values[0]![0]
const rowsOf = async (s: LocalSource) => (await s.book()).sheets[0]!.rows ?? {}

describe('localSource: the outline of a sheet', () => {
  it("returns the groups from the levels and hidden rows or columns, with Excel's summaries by default; a sheet without groups has none", async () => {
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

  it('accepts the summary settings and treats a height of 0 as hidden', async () => {
    const { s } = await opened(book({ rows: { '3:4': { outlineLevel: 1, height: 0 } }, columns: {}, outline: { summaryBelow: false } }))
    expect(await outlineOf(s)).toEqual({ rows: [{ from: 3, to: 4, level: 1, collapsed: true }], columns: [], summaryBelow: false, summaryRight: true })
  })
})

describe('localSource: outline commands', () => {
  it('Show Detail of the outer group keeps the inner one collapsed; each command is a single undo step, sent as a reload', async () => {
    const { s, events, run } = await opened()
    expect(await hiddenRows(s)).toEqual([6, 7, 8, 9])
    expect(await run({ op: 'collapse', rows: '2:10', level: 1 })).toEqual({ ok: true, changes: [expect.objectContaining({ cell: 'C12' })] })
    expect(events).toEqual([{ kind: 'reload' }])
    expect(await hiddenRows(s)).toEqual([2, 3, 4, 5, 6, 7, 8, 9, 10])
    expect(await subtotal(s)).toBe(1 + 11)
    expect((await rowsOf(s))['11']).toEqual({ collapsed: true })
    await run({ op: 'expand', rows: '2:10', level: 1 })
    expect(await hiddenRows(s)).toEqual([6, 7, 8, 9])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
    await s.undo()
    expect(await hiddenRows(s)).toEqual([2, 3, 4, 5, 6, 7, 8, 9, 10])
  })

  it('the button of a group that shares its summary row with the outer group acts on all the detail of that row', async () => {
    // rows 3:7 at level 1 with 5:7 at level 2: both groups end at row 7 and have their buttons on row 8
    const { s, run } = await opened(book({ rows: { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 } }, columns: {} }))
    expect(await run({ op: 'collapse', rows: '5:7', level: 2 })).toMatchObject({ ok: true })
    expect(await hiddenRows(s)).toEqual([3, 4, 5, 6, 7])
    expect((await outlineOf(s))!.rows).toEqual([
      { from: 3, to: 7, level: 1, collapsed: true },
      { from: 5, to: 7, level: 2, collapsed: true },
    ])
    expect(await run({ op: 'expand', rows: '5:7', level: 2 })).toMatchObject({ ok: true })
    expect(await hiddenRows(s)).toEqual([])
    expect(await rowsOf(s)).toEqual({ '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 } })
  })

  it('rows hidden manually under a summary row flag it at once: saveXlsx writes the flag, and the group stays collapsed inside the outer one', async () => {
    const { s, run } = await opened(book({ rows: NESTED, columns: {} }))
    await s.layout!({ sheet: 'S', rows: '5:7', hidden: true })
    expect((await outlineOf(s))!.rows).toEqual([
      { from: 3, to: 10, level: 1, collapsed: false },
      { from: 5, to: 7, level: 2, collapsed: true },
    ])
    const read = await openXlsx((await saveXlsx(s)).data)
    expect(read.book.sheets[0]!.rows!['8']).toEqual({ outlineLevel: 1, collapsed: true })
    await run({ op: 'collapse', rows: '3:10', level: 1 })
    expect(await hiddenRows(s)).toEqual([3, 4, 5, 6, 7, 8, 9, 10])
    await run({ op: 'expand', rows: '3:10', level: 1 })
    expect(await hiddenRows(s)).toEqual([5, 6, 7])
  })

  it('a level button shows manually hidden grouped rows; a summary row it hides keeps its flag, so its group stays collapsed', async () => {
    const { s, run } = await opened(book({ rows: { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 }, '8': { outlineLevel: 1 }, '9': { outlineLevel: 1, hidden: true }, '10': { outlineLevel: 1 } }, columns: {} }))
    expect(await hiddenRows(s)).toEqual([9])
    await run({ op: 'showLevel', axis: 'rows', level: 2 })
    expect(await hiddenRows(s)).toEqual([5, 6, 7])
    expect((await rowsOf(s))['8']).toEqual({ outlineLevel: 1, collapsed: true })
    await run({ op: 'showLevel', axis: 'rows', level: 1 })
    expect(await hiddenRows(s)).toEqual([3, 4, 5, 6, 7, 8, 9, 10])
    expect((await rowsOf(s))['8']).toEqual({ outlineLevel: 1, hidden: true, collapsed: true })
    await run({ op: 'expand', rows: '3:10', level: 1 })
    expect(await hiddenRows(s)).toEqual([5, 6, 7])
    await s.layout!({ sheet: 'S', rows: '4', hidden: true })
    await run({ op: 'showLevel', axis: 'rows', level: 3 })
    expect(await hiddenRows(s)).toEqual([])
    expect(Object.values(await rowsOf(s)).some((r) => r.collapsed)).toBe(false)
  })

  it("Group and Ungroup over columns; a level button over rows; a rejection is the library's message", async () => {
    const { s, run } = await opened()
    expect(await run({ op: 'group', columns: 'C:C' })).toMatchObject({ ok: true })
    expect((await outlineOf(s))!.columns).toEqual([
      { from: 2, to: 4, level: 1, collapsed: false },
      { from: 3, to: 3, level: 2, collapsed: false },
    ])
    await run({ op: 'ungroup', columns: 'B:D' })
    expect((await outlineOf(s))!.columns).toEqual([{ from: 3, to: 3, level: 1, collapsed: false }])
    expect(await run({ op: 'ungroup', columns: 'F:G' })).toEqual({ ok: false, reason: 'refused', message: 'the columns F:G are not in a group' })
    await run({ op: 'showLevel', axis: 'rows', level: 3 })
    expect(await hiddenRows(s)).toEqual([])
    expect(await run({ op: 'collapse', rows: '3:4', level: 2 })).toEqual({ ok: false, reason: 'refused', message: 'No group of level 2 over the rows 3:4' })
  })

  it("Group beyond the seventh level and Ungroup of rows in no group are rejected in the workbook's language; Group next to the seventh level deepens the rest", async () => {
    const { s, run } = await opened(book({ rows: {}, columns: {} }, 'ru-RU'))
    for (let i = 0; i < 7; i++) expect(await run({ op: 'group', rows: '3:6' })).toMatchObject({ ok: true })
    expect(await run({ op: 'group', rows: '3:6' })).toEqual({ ok: false, reason: 'refused', message: 'Нельзя сгруппировать строки 3:6: они уже на самом глубоком уровне структуры, 7' })
    expect(await run({ op: 'ungroup', rows: '8:9' })).toEqual({ ok: false, reason: 'refused', message: 'Строки 8:9 не входят ни в одну группу' })
    expect(await run({ op: 'group', rows: '2:6' })).toMatchObject({ ok: true })
    expect(await rowsOf(s)).toEqual({ '2': { outlineLevel: 1 }, '3:6': { outlineLevel: 7 } })
  })

  it('undo reverts a command as a whole, the hidden rows and the flags, and redo applies it again', async () => {
    const { s, run } = await opened(book({ rows: NESTED, columns: {} }))
    await run({ op: 'collapse', rows: '5:7', level: 2 })
    const collapsed = await rowsOf(s)
    expect(collapsed).toEqual({ '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2, hidden: true }, '8': { outlineLevel: 1, collapsed: true }, '9:10': { outlineLevel: 1 } })
    expect(await s.undo()).toMatchObject({ ok: true })
    expect(await rowsOf(s)).toEqual(NESTED)
    expect(await hiddenRows(s)).toEqual([])
    expect((await outlineOf(s))!.rows[1]).toEqual({ from: 5, to: 7, level: 2, collapsed: false })
    await s.redo()
    expect(await rowsOf(s)).toEqual(collapsed)
    expect(await hiddenRows(s)).toEqual([5, 6, 7])
  })

  it('a group at the edge of the sheet has no summary row, so no button: its command is rejected, and a level button hides it', async () => {
    const { s, run } = await opened(book({ rows: { '1:3': { outlineLevel: 1 } }, columns: {}, outline: { summaryBelow: false } }))
    expect(await run({ op: 'collapse', rows: '1:3', level: 1 })).toEqual({ ok: false, reason: 'refused', message: "The group of rows 1:3 has no summary row within the sheet: Excel draws no button for it" })
    expect(await hiddenRows(s)).toEqual([])
    await run({ op: 'showLevel', axis: 'rows', level: 1 })
    expect(await hiddenRows(s)).toEqual([1, 2, 3])
    expect((await outlineOf(s))!.rows).toEqual([{ from: 1, to: 3, level: 1, collapsed: true }])
  })

  it('a row hidden by a height of 0 is unhidden at the default height', async () => {
    const { s, run } = await opened(book({ rows: { '3:4': { outlineLevel: 1, height: 0 } }, columns: {} }))
    await run({ op: 'expand', rows: '3:4', level: 1 })
    expect(await hiddenRows(s)).toEqual([])
    expect(await rowsOf(s)).toEqual({ '3:4': { outlineLevel: 1 } })
  })

  it('an unknown sheet fails', async () => {
    const { s } = await opened()
    expect(await s.outline({ op: 'showLevel', sheet: 'Nope', axis: 'rows', level: 1 })).toMatchObject({ ok: false, reason: 'failed' })
  })
})
