// @vitest-environment node
// Turned text over `./local` (xlcalc 0.67): rows fitted to it after typing and after Format Cells, the sheet told to read
// them again, and a turned number shown whole in a narrow column.
import { afterEach, describe, expect, it } from 'vitest'
import type { SourceEvent } from '../src/contract/source'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

const turned = (cells: Record<string, { value: string | number; style?: number }>) =>
  source({ sheets: [{ name: 'S', cells, columns: { A: { width: 3 } } }], styles: [{}, { alignment: { rotation: 90 } }] })

describe('turned text in ./local', () => {
  it('grows the row of a turned cell typed into and reloads the sheet', async () => {
    const s = turned({ B1: { value: 'x', style: 1 } })
    const events: SourceEvent['kind'][] = []
    s.subscribe((e) => events.push(e.kind))
    expect((await win(s, 'S', 1, 2, 1, 2)).rowHeights[1]).toBeUndefined()
    await s.edit([{ sheet: 'S', cell: 'B1', text: 'a long turned text here' }])
    expect(events).toContain('reload')
    expect((await win(s, 'S', 1, 2, 1, 2)).rowHeights[1]).toBeGreaterThan(60)
  })

  it('fits the row to text turned through Format Cells and shows a turned number whole in a narrow column', async () => {
    const s = turned({ A1: { value: 1234567890 }, B1: { value: 'a long level text here' } })
    expect((await win(s, 'S', 1, 1, 1, 2)).text[0]![0]).toBe('##')
    const events: SourceEvent['kind'][] = []
    s.subscribe((e) => events.push(e.kind))
    await s.format('S', [{ range: 'A1:B1', style: { alignment: { rotation: 45 } } }])
    expect(events).toContain('reload')
    const w = await win(s, 'S', 1, 1, 1, 2)
    expect(w.rowHeights[1]).toBeGreaterThan(40)
    expect(w.text[0]![0]).toBe('1234567890')
  })
})
