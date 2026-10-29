// @vitest-environment node
// localSource and the sheet's look beside its cells: tab colors through `setLayout`, a single undo step, kept in the workbook.
import { afterEach, describe, expect, it } from 'vitest'
import type { SourceEvent } from '../src/contract/source'
import { closeAll, source } from './local-fixture'

afterEach(closeAll)

const two = () => source({ sheets: [{ name: 'A', cells: { A1: { value: 1 } }, tabColor: '#00B050' }, { name: 'B', cells: {} }] })

describe('localSource: tab colors', () => {
  it("lists a sheet's tab color from the workbook", async () => {
    const list = await two().sheets()
    expect(list[0]!.tabColor).toBe('#00B050')
    expect(list[1]!.tabColor).toBeUndefined()
  })

  it('sets and removes a tab color in one undo step each, keeps it in the workbook, and tells subscribers to read the sheets again', async () => {
    const s = two()
    const events: SourceEvent['kind'][] = []
    s.subscribe((e) => events.push(e.kind))
    expect(await s.tabColor('B', '#FF0000')).toMatchObject({ ok: true })
    expect(events).toContain('reload')
    expect((await s.sheets())[1]!.tabColor).toBe('#FF0000')
    expect((await s.book()).sheets[1]!.tabColor).toBe('#FF0000')
    expect(s.history.undo).toBe(1)
    expect(await s.tabColor('A', null)).toMatchObject({ ok: true })
    expect((await s.sheets())[0]!.tabColor).toBeUndefined()
    await s.undo()
    expect((await s.sheets())[0]!.tabColor).toBe('#00B050')
    await s.undo()
    expect((await s.sheets())[1]!.tabColor).toBeUndefined()
  })
})
