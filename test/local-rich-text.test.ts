// @vitest-environment node
import { afterEach, describe, expect, it } from 'vitest'
import { openXlsx, saveXlsx } from '../src/local'
import type { SheetWindow } from '../src/contract/source'
import { closeAll, source } from './local-fixture'

afterEach(closeAll)
const runs = [{ start: 0, font: { bold: true } }, { start: 2, font: { italic: true, color: '#FF0000' } }]
const book = () => ({ sheets: [{ name: 'S', cells: { A1: { value: '😀text', richText: structuredClone(runs) } } }] })
const windowOf = async (s: ReturnType<typeof source>) => await s.window({ sheet: 'S', row: 1, rows: 1, col: 1, cols: 1 }) as SheetWindow

describe('localSource: rich text', () => {
  it('carries run fonts in windows and clears and restores them in cell updates', async () => {
    const s = source(book())
    expect((await windowOf(s)).richText?.A1).toEqual(runs)
    const changed = await s.edit([{ sheet: 'S', cell: 'A1', text: 'plain' }])
    expect(changed).toMatchObject({ ok: true, changes: [expect.objectContaining({ cell: 'A1', richText: null })] })
    expect((await windowOf(s)).richText?.A1).toBeUndefined()
    await s.undo()
    expect((await windowOf(s)).richText?.A1).toEqual(runs)
  })

  it('preserves runs through font patches and both xlsx writers', async () => {
    const original = (await saveXlsx(source(book()))).data
    const kept = await openXlsx(original, { keep: true })
    const s = source(kept.book, { session: { cached: kept.cached } })
    await s.format('S', [{ range: 'A1', style: { font: { underline: 'double' } } }])
    expect((await windowOf(s)).richText?.A1).toMatchObject(runs.map(r => ({ ...r, font: { ...r.font, underline: 'double' } })))
    for (const opts of [{}, { over: original }]) {
      const opened = await openXlsx((await saveXlsx(s, opts)).data)
      expect(opened.book.sheets[0]!.cells.A1!.richText).toMatchObject(runs.map(r => ({ ...r, font: { ...r.font, underline: 'double' } })))
    }
  })
})
