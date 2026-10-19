// @vitest-environment node
// Excel's outline commands of memorySource (src/memory-outline.ts, xlcalc's `outlineEdit`) against what Excel for Windows 16
// (build 20430) did over COM on a synthetic book, 12.10.2026: each case is a measured scenario, the hidden rows and the
// collapsed marks as Excel saved them; and the cases xlcalc's later measurements added, as ./local gives them
// (test/local-outline.test.ts).
import { outlineGroups } from 'xlcalc/a1'
import { describe, expect, it } from 'vitest'
import { memorySource } from '../src'
import { outlineCommand, outlineSheet, type AxisLines } from '../src/memory-outline'
import type { OutlineChange, SheetWindow } from '../src/contract/source'

/** A command without its sheet, each kind of `OutlineChange` kept apart. */
type Change = OutlineChange extends infer C ? (C extends unknown ? Omit<C, 'sheet'> : never) : never

/** A sheet's rows one by one, as Excel's rows: level, hidden, the collapsed mark. */
class Rows implements AxisLines {
  levels = new Map<number, number>()
  hidden = new Set<number>()
  marks = new Set<number>()
  constructor(levels: Record<number, number> = {}) {
    for (const [r, l] of Object.entries(levels)) this.levels.set(Number(r), l)
  }
  /** Apply a command as memorySource does; its refusal's message, or null. */
  run(change: Change, after = true): string | null {
    return outlineCommand(this, { summaryBelow: after, summaryRight: after }, { sheet: 'S', ...change } as OutlineChange)
  }
  /** The groups xlcalc draws from the rows, summaries below. */
  groups() {
    return outlineGroups(outlineSheet({ rows: this, columns: new Rows() })).rows
  }
  /** Rows 1..13 as the measurement printed them: `3:1h` is row 3 at level 1, hidden. */
  state(n = 13): string {
    return Array.from({ length: n }, (_, i) => `${i + 1}:${this.levels.get(i + 1) ?? 0}${this.hidden.has(i + 1) ? 'h' : ''}`).join(' ')
  }
  get collapsed(): number[] {
    return [...this.marks].sort((a, b) => a - b)
  }
}
const lvl = (r: string, level: number) => ({ rows: r, level })

describe('groups as Excel draws them', () => {
  it('a level is the longest runs of lines at it or deeper; outer before inner; collapsed once every line under its summary is hidden, by hand too', () => {
    const rows = new Rows({ 3: 1, 4: 1, 5: 2, 6: 2, 7: 2, 8: 1, 9: 1, 10: 1, 12: 1 })
    for (const r of [5, 6, 7, 12]) rows.hidden.add(r)
    rows.marks.add(8)
    expect(rows.groups()).toEqual([
      { first: 3, last: 10, level: 1, summary: 11, collapsed: false },
      { first: 5, last: 7, level: 2, summary: 8, collapsed: true },
      { first: 12, last: 12, level: 1, summary: 13, collapsed: true },
    ])
  })
  it('a group whose summary line is hidden inside a hidden group is collapsed by the mark it kept, as Excel keeps it', () => {
    const hidden = (marked: boolean) => {
      const rows = new Rows({ 3: 1, 4: 1, 5: 2, 6: 2, 7: 2, 8: 1, 9: 1, 10: 1 })
      for (let r = 3; r <= 10; r++) rows.hidden.add(r)
      rows.marks.add(11)
      if (marked) rows.marks.add(8)
      return rows.groups().map((g) => [g.first, g.last, g.collapsed])
    }
    expect(hidden(false)).toEqual([[3, 10, true], [5, 7, false]])
    expect(hidden(true)).toEqual([[3, 10, true], [5, 7, true]])
  })
})

describe('Group and Ungroup', () => {
  it('S1: each Group puts the lines a level deeper; the eighth is refused, the levels stay at 7', () => {
    const rows = new Rows()
    for (let i = 1; i <= 7; i++) expect(rows.run({ op: 'group', rows: '3:6' })).toBeNull()
    expect(rows.run({ op: 'group', rows: '3:6' })).toBe('the rows 3:6 are at the deepest outline level, 7, and are not grouped further')
    expect(rows.state(7)).toBe('1:0 2:0 3:7 4:7 5:7 6:7 7:0')
    // over lines at 7 and others Excel deepens the others and leaves those at 7 (xlcalc's docs/outline.md)
    expect(rows.run({ op: 'group', rows: '2:6' })).toBeNull()
    expect(rows.state(7)).toBe('1:0 2:1 3:7 4:7 5:7 6:7 7:0')
  })

  it('S2: Ungroup of lines in no group is refused; over a mix it lowers the grouped ones; over a part of a group that part', () => {
    const rows = new Rows()
    expect(rows.run({ op: 'ungroup', rows: '3:6' })).toBe('the rows 3:6 are in no group')
    rows.run({ op: 'group', rows: '4:5' })
    expect(rows.run({ op: 'ungroup', rows: '3:6' })).toBeNull()
    expect(rows.state(7)).toBe('1:0 2:0 3:0 4:0 5:0 6:0 7:0')
    rows.run({ op: 'group', rows: '4:5' })
    rows.run({ op: 'ungroup', rows: '5:5' })
    expect(rows.state(7)).toBe('1:0 2:0 3:0 4:1 5:0 6:0 7:0')
  })

  it('S8, S9: Ungroup leaves hidden lines hidden and the summary line its mark', () => {
    const rows = new Rows({ 3: 1, 4: 1, 5: 1, 6: 1 })
    rows.run({ op: 'collapse', ...lvl('3:6', 1) })
    rows.run({ op: 'ungroup', rows: '3:6' })
    expect(rows.state(7)).toBe('1:0 2:0 3:0h 4:0h 5:0h 6:0h 7:0')
    expect(rows.collapsed).toEqual([7])
    const nested = new Rows({ 3: 1, 4: 1, 5: 2, 6: 2, 7: 2, 8: 1, 9: 1, 10: 1 })
    nested.run({ op: 'collapse', ...lvl('5:7', 2) })
    nested.run({ op: 'ungroup', rows: '3:10' })
    expect(nested.state(11)).toBe('1:0 2:0 3:0 4:0 5:1h 6:1h 7:1h 8:0 9:0 10:0 11:0')
    expect(nested.collapsed).toEqual([8])
  })
})

describe('a group’s button', () => {
  const nested = () => new Rows({ 3: 1, 4: 1, 5: 2, 6: 2, 7: 2, 8: 1, 9: 1, 10: 1 })

  it('S3: a group collapsed inside a collapsed one stays collapsed when the outer is expanded, by its mark', () => {
    const rows = nested()
    rows.run({ op: 'collapse', ...lvl('5:7', 2) })
    expect(rows.state(11)).toBe('1:0 2:0 3:1 4:1 5:2h 6:2h 7:2h 8:1 9:1 10:1 11:0')
    rows.run({ op: 'collapse', ...lvl('3:10', 1) })
    expect(rows.state(11)).toBe('1:0 2:0 3:1h 4:1h 5:2h 6:2h 7:2h 8:1h 9:1h 10:1h 11:0')
    expect(rows.collapsed).toEqual([8, 11])
    rows.run({ op: 'expand', ...lvl('3:10', 1) })
    expect(rows.state(11)).toBe('1:0 2:0 3:1 4:1 5:2h 6:2h 7:2h 8:1 9:1 10:1 11:0')
    expect(rows.collapsed).toEqual([8])
  })

  it('S3b: one that was expanded comes back expanded', () => {
    const rows = nested()
    rows.run({ op: 'collapse', ...lvl('3:10', 1) })
    expect(rows.collapsed).toEqual([11])
    rows.run({ op: 'expand', ...lvl('3:10', 1) })
    expect(rows.state(11)).toBe('1:0 2:0 3:1 4:1 5:2 6:2 7:2 8:1 9:1 10:1 11:0')
  })

  it('S4: groups that share their summary line open together', () => {
    const rows = new Rows({ 3: 1, 4: 2, 5: 2, 6: 2 })
    rows.run({ op: 'collapse', ...lvl('3:6', 1) })
    expect(rows.collapsed).toEqual([7])
    rows.run({ op: 'expand', ...lvl('3:6', 1) })
    expect(rows.state(7)).toBe('1:0 2:0 3:1 4:2 5:2 6:2 7:0')
    expect(rows.collapsed).toEqual([])
  })

  it('S5, S6: Show Detail shows a line hidden by hand; a group hidden by hand is collapsed', () => {
    const rows = new Rows({ 3: 1, 4: 1, 5: 1, 6: 1 })
    rows.hidden.add(4)
    rows.run({ op: 'collapse', ...lvl('3:6', 1) })
    rows.run({ op: 'expand', ...lvl('3:6', 1) })
    expect(rows.state(7)).toBe('1:0 2:0 3:1 4:1 5:1 6:1 7:0')
    for (const r of [3, 4, 5, 6]) rows.hidden.add(r)
    expect(rows.groups()).toEqual([{ first: 3, last: 6, level: 1, summary: 7, collapsed: true }])
    rows.run({ op: 'expand', ...lvl('3:6', 1) })
    expect(rows.state(7)).toBe('1:0 2:0 3:1 4:1 5:1 6:1 7:0')
  })

  it('S10: with summaries above, the mark goes on the line before the group', () => {
    const rows = new Rows({ 3: 1, 4: 1, 5: 1, 6: 1 })
    rows.run({ op: 'collapse', ...lvl('3:6', 1) }, false)
    expect([rows.state(7), rows.collapsed]).toEqual(['1:0 2:0 3:1h 4:1h 5:1h 6:1h 7:0', [2]])
    rows.run({ op: 'expand', ...lvl('3:6', 1) }, false)
    expect([rows.state(7), rows.collapsed]).toEqual(['1:0 2:0 3:1 4:1 5:1 6:1 7:0', []])
  })

  it('names the group by its lines and level; another is refused', () => {
    expect(nested().run({ op: 'collapse', ...lvl('3:9', 1) })).toMatch(/No group of level 1/)
    expect(nested().run({ op: 'collapse', ...lvl('5:7', 1) })).toMatch(/No group of level 1/)
  })
})

describe('a level button', () => {
  it('S7: level k hides the lines of level k and deeper, shows the others; lines outside groups stay; marks on the summaries left in view', () => {
    const rows = new Rows({ 3: 1, 4: 1, 5: 2, 6: 2, 7: 2, 8: 1, 9: 1, 10: 1 })
    rows.hidden.add(9)
    rows.hidden.add(13)
    rows.run({ op: 'showLevel', axis: 'rows', level: 1 })
    expect([rows.state(), rows.collapsed]).toEqual(['1:0 2:0 3:1h 4:1h 5:2h 6:2h 7:2h 8:1h 9:1h 10:1h 11:0 12:0 13:0h', [11]])
    rows.run({ op: 'showLevel', axis: 'rows', level: 2 })
    expect([rows.state(), rows.collapsed]).toEqual(['1:0 2:0 3:1 4:1 5:2h 6:2h 7:2h 8:1 9:1 10:1 11:0 12:0 13:0h', [8]])
    rows.run({ op: 'showLevel', axis: 'rows', level: 3 })
    expect([rows.state(), rows.collapsed]).toEqual(['1:0 2:0 3:1 4:1 5:2 6:2 7:2 8:1 9:1 10:1 11:0 12:0 13:0h', []])
    rows.run({ op: 'showLevel', axis: 'rows', level: 1 })
    rows.run({ op: 'expand', ...lvl('3:10', 1) })
    expect(rows.state()).toBe('1:0 2:0 3:1 4:1 5:2 6:2 7:2 8:1 9:1 10:1 11:0 12:0 13:0h')
  })

  it('S11: columns go by the same rules', () => {
    const cols = new Rows({ 3: 1, 4: 2, 5: 1 })
    const run = (op: 'collapse' | 'expand', columns: string, level: number) => cols.run({ op, columns, level })
    run('collapse', 'D', 2)
    run('collapse', 'C:E', 1)
    run('expand', 'C:E', 1)
    expect(cols.state(6)).toBe('1:0 2:0 3:1 4:2h 5:1 6:0')
  })
})

describe('memorySource: the commands as Excel and ./local give them', () => {
  /** A memory sheet with these row levels and options, its commands and what it shows. */
  function sheet(rows: Record<string, number>, opts: { heights?: Record<string, number>; summaryBelow?: boolean } = {}) {
    const src = memorySource({ sheets: [{ name: 'S', cells: { A1: 1 }, heights: opts.heights, outline: { rows, ...(opts.summaryBelow !== undefined ? { summaryBelow: opts.summaryBelow } : {}) } }] })
    return {
      run: (change: Change) => src.outline!({ sheet: 'S', ...change } as OutlineChange),
      hidden: async () => Object.entries(((await src.window({ sheet: 'S', row: 1, rows: 12, col: 1, cols: 1 })) as SheetWindow).rowHeights).filter(([, h]) => h === 0).map(([r]) => Number(r)),
      heights: async () => ((await src.window({ sheet: 'S', row: 1, rows: 12, col: 1, cols: 1 })) as SheetWindow).rowHeights,
      groups: async () => (await src.sheets())[0]!.outline?.rows,
    }
  }

  it('the button of a group sharing its summary row with the outer group acts on the whole detail of that row', async () => {
    const s = sheet({ '3:4': 1, '5:7': 2 })
    expect(await s.run({ op: 'collapse', ...lvl('5:7', 2) })).toEqual({ ok: true, changes: [] })
    expect(await s.hidden()).toEqual([3, 4, 5, 6, 7])
    expect(await s.groups()).toEqual([
      { from: 3, to: 7, level: 1, collapsed: true },
      { from: 5, to: 7, level: 2, collapsed: true },
    ])
    await s.run({ op: 'expand', ...lvl('5:7', 2) })
    expect(await s.hidden()).toEqual([])
  })

  it('a button is refused when its group already shows as asked', async () => {
    const s = sheet({ '3:4': 1, '5:7': 2, '8:10': 1 })
    expect(await s.run({ op: 'expand', ...lvl('5:7', 2) })).toEqual({ ok: false, reason: 'refused', message: 'the detail of row 8 is shown already' })
    await s.run({ op: 'collapse', ...lvl('5:7', 2) })
    expect(await s.run({ op: 'collapse', ...lvl('5:7', 2) })).toEqual({ ok: false, reason: 'refused', message: 'the detail of row 8 is hidden already' })
    expect(await s.hidden()).toEqual([5, 6, 7])
  })

  it('a group at the sheet’s edge has no summary row, so no button: its command is refused, a level button hides it', async () => {
    const s = sheet({ '1:3': 1 }, { summaryBelow: false })
    expect(await s.run({ op: 'collapse', ...lvl('1:3', 1) })).toEqual({ ok: false, reason: 'refused', message: "The group of rows 1:3 has no summary row within the sheet: Excel draws no button for it" })
    expect(await s.hidden()).toEqual([])
    await s.run({ op: 'showLevel', axis: 'rows', level: 1 })
    expect(await s.hidden()).toEqual([1, 2, 3])
    expect(await s.groups()).toEqual([{ from: 1, to: 3, level: 1, collapsed: true }])
  })

  it('a line hidden by a height of 0 is shown again at the default height; a level button past 8 fails', async () => {
    const s = sheet({ '3:4': 1 }, { heights: { '3:4': 0, '6': 30 } })
    await s.run({ op: 'expand', ...lvl('3:4', 1) })
    expect(await s.heights()).toEqual({ 6: 30 })
    expect(await s.run({ op: 'showLevel', axis: 'rows', level: 9 })).toMatchObject({ ok: false, reason: 'failed' })
  })
})
