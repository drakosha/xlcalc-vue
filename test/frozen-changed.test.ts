// A source's `changed` over a sheet with frozen panes: every drawn cell shows it at once, in whatever order the windows arrived.
import { afterEach, describe, expect, it } from 'vitest'
import { cellAddress } from 'xlcalc/a1'
import type { SheetInfo, WindowRequest } from '../src/contract/source'
import { localSource } from '../src/local'
import { cellText, fakeSource, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
})

type Fake = ReturnType<typeof fakeSource>
type Freeze = NonNullable<SheetInfo['freeze']>

/** The window a request is for, by the frozen rows and columns it covers. */
function kindOf(q: WindowRequest, freeze?: Freeze): string {
  const top = !!freeze?.rows && q.row === 1 && q.rows === freeze.rows
  const left = !!freeze?.cols && q.col === 1 && q.cols === freeze.cols
  return top && left ? 'corner' : top ? 'top' : left ? 'left' : 'body'
}

/** Answers the held requests one by one in the order of `kinds`, drawing after each, then whatever those responses asked. */
async function answerIn(f: Fake, freeze: Freeze | undefined, kinds: string[]): Promise<void> {
  for (const kind of kinds) {
    const held = f.held.find((p) => kindOf(p.req, freeze) === kind)
    if (!held) continue
    f.release(held)
    await settle()
  }
  while (f.held.length) {
    f.release(f.held[0]!)
    await settle()
  }
}

const sizes = [{ maxRow: 100, maxCol: 10 }, { maxRow: 6, maxCol: 3 }]
const freezes: (Freeze | undefined)[] = [{ rows: 1, cols: 0 }, { rows: 0, cols: 1 }, { rows: 1, cols: 1 }, { rows: 2, cols: 0 }, undefined]
const orders = [
  ['body', 'top', 'left', 'corner'],
  ['top', 'left', 'corner', 'body'],
  ['left', 'body', 'corner', 'top'],
  ['body', 'corner', 'left', 'top'],
]
const events = [['A1', 'A2'], ['C2'], ['C5']]
const at = (cell: string): [number, number] => [Number(cell.slice(1)), cell.charCodeAt(0) - 64]

describe('changed over frozen panes', () => {
  for (const size of sizes) {
    for (const freeze of freezes) {
      for (const kinds of orders) {
        for (const cells of events) {
          it(`${size.maxRow} rows, freeze ${JSON.stringify(freeze ?? null)}, windows ${kinds.join(' ')}, changing ${cells.join(' ')}`, async () => {
            const sheets: SheetInfo[] = [{ name: 'S', ...size, ...(freeze ? { freeze } : {}) }]
            const f = fakeSource(sheets, { manual: true, cell: (_, r, c) => `old ${cellAddress(r, c)}` })
            // windows as a real source's: the same sizes in each, and `tables`, which ends the search for them at the first window
            const read = f.source.window.bind(f.source)
            f.source.window = async (req, signal) => ({ ...(await read(req, signal)), widths: new Array(req.cols).fill(80), tables: [] })
            const w = await mountSheet(f.source, { readOnly: true })
            await answerIn(f, freeze, kinds)
            for (const cell of cells) expect(cellText(w, ...at(cell))).toBe(`old ${cell}`)
            f.emit({ kind: 'changed', changes: cells.map((cell) => ({ sheet: 'S', cell, text: `new ${cell}`, value: `new ${cell}` })) })
            await settle()
            expect(cells.map((cell) => cellText(w, ...at(cell)))).toEqual(cells.map((cell) => `new ${cell}`))
          })
        }
      }
    }
  }
})

describe('changed over frozen rows from a real source', () => {
  it('localSource: an edit reaches the formulas that depend on it in the frozen row and in the body, unselected', async () => {
    const s = localSource({ sheets: [{ name: 'S', cells: { A1: { formula: 'C5+1' }, A2: { formula: 'C5*2' }, C5: { value: 1 } }, frozen: { rows: 1 } }] }, { calc: 'inline' })
    const w = await mountSheet(s, { readOnly: true })
    await settle()
    expect([cellText(w, 1, 1), cellText(w, 2, 1)]).toEqual(['2', '2'])
    expect(await s.edit([{ sheet: 'S', cell: 'C5', text: '7' }])).toMatchObject({ ok: true })
    await settle()
    expect([cellText(w, 1, 1), cellText(w, 2, 1), cellText(w, 5, 3)]).toEqual(['8', '14', '7'])
    s.close()
  })
})
