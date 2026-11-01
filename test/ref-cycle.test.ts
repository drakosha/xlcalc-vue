// Excel's F4 in an edited formula, as Excel for Mac 16.114 cycled references with Cmd+T (its F4), measured 31.10 from
// AppleScript: four presses on each formula, the caret or the selection set before the first.
import { describe, expect, it } from 'vitest'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { cycleRefs } from '../src/ref-cycle'

/** The texts after one to four presses, each press on the previous result's caret or selection. */
function presses(text: string, from = text.length, to = from, opts = {}): string[] {
  const out: string[] = []
  let at = { text, from, to }
  for (let i = 0; i < 4; i++) {
    at = cycleRefs(at.text, at.from, at.to, opts) ?? at
    out.push(at.text)
  }
  return out
}

describe('F4 at a caret', () => {
  it('cycles a cell A1, $A$1, A$1, $A1, A1, with the caret after it or inside it', () => {
    expect(presses('=A1')).toEqual(['=$A$1', '=A$1', '=$A1', '=A1'])
    expect(presses('=A1+1', 3)).toEqual(['=$A$1+1', '=A$1+1', '=$A1+1', '=A1+1'])
    expect(presses('=A1+B2', 4).slice(0, 2)).toEqual(['=A1+$B$2', '=A1+B$2'])
    expect(presses('=A1+B2', 3).slice(0, 2)).toEqual(['=$A$1+B2', '=A$1+B2'])
    expect(presses('=SUM(A1,5)', 7).slice(0, 2)).toEqual(['=SUM($A$1,5)', '=SUM(A$1,5)'])
  })

  it("in a typed range only the cell at the caret, its own state; a trim reference's points stay", () => {
    expect(presses('=A1:B2')).toEqual(['=A1:$B$2', '=A1:B$2', '=A1:$B2', '=A1:B2'])
    expect(presses('=A1:B2', 4)).toEqual(['=A1:$B$2', '=A1:B$2', '=A1:$B2', '=A1:B2'])
    expect(presses('=A1:B2', 3)).toEqual(['=$A$1:B2', '=A$1:B2', '=$A1:B2', '=A1:B2'])
    expect(presses('=A1:B2', 1)).toEqual(['=$A$1:B2', '=A$1:B2', '=$A1:B2', '=A1:B2'])
    expect(presses('=$A1:B2')).toEqual(['=$A1:$B$2', '=$A1:B$2', '=$A1:$B2', '=$A1:B2'])
    expect(presses('=A$1:$B2')).toEqual(['=A$1:B2', '=A$1:$B$2', '=A$1:B$2', '=A$1:$B2'])
    expect(presses('=A1:.B2')).toEqual(['=A1:.$B$2', '=A1:.B$2', '=A1:.$B2', '=A1:.B2'])
    expect(presses('=A1.:.B2')).toEqual(['=A1.:.$B$2', '=A1.:.B$2', '=A1.:.$B2', '=A1.:.B2'])
  })

  it("columns and rows whole, between relative and absolute, from the first end's state", () => {
    expect(presses('=A:B')).toEqual(['=$A:$B', '=A:B', '=$A:$B', '=A:B'])
    expect(presses('=$A:B')).toEqual(['=A:B', '=$A:$B', '=A:B', '=$A:$B'])
    expect(presses('=1:2')).toEqual(['=$1:$2', '=1:2', '=$1:$2', '=1:2'])
    expect(presses('=$3:4').slice(0, 1)).toEqual(['=3:4'])
  })

  it("keeps a sheet's name, a spill's # and the locale's separators; nothing outside a reference, in a table reference or in text that is no formula", () => {
    expect(presses("=SUM('My sheet'!B2)", 17).slice(0, 1)).toEqual(["=SUM('My sheet'!$B$2)"])
    expect(presses('=C5#').slice(0, 1)).toEqual(['=$C$5#'])
    expect(presses('=СУММ(A1;1,5)', 8, 8, { locale: ruRU }).slice(0, 1)).toEqual(['=СУММ($A$1;1,5)'])
    expect(cycleRefs('=SUM(A1,5)', 9, 9, {})).toBeNull()
    expect(cycleRefs('=1+2', 4, 4, {})).toBeNull()
    expect(cycleRefs('=Sales[Qty]', 11, 11, {})).toBeNull()
    expect(cycleRefs('A1', 2, 2, {})).toBeNull()
  })

  it('leaves the caret at the end of what changed', () => {
    expect(cycleRefs('=A1:B2', 3, 3, {})).toEqual({ text: '=$A$1:B2', from: 5, to: 5 })
    expect(cycleRefs('=$A$1+B2', 5, 5, {})).toEqual({ text: '=A$1+B2', from: 4, to: 4 })
  })
})

describe('F4 over a selection', () => {
  it('cycles every reference it touches, each from its own state, a range whole from its first end; the selection follows', () => {
    expect(presses('=A1:B2', 0, 6)).toEqual(['=$A$1:$B$2', '=A$1:B$2', '=$A1:$B2', '=A1:B2'])
    expect(presses('=SUM(A1,B2)+C3', 0, 14)).toEqual(['=SUM($A$1,$B$2)+$C$3', '=SUM(A$1,B$2)+C$3', '=SUM($A1,$B2)+$C3', '=SUM(A1,B2)+C3'])
    expect(presses('=$A1:B2', 0, 7)).toEqual(['=A1:B2', '=$A$1:$B$2', '=A$1:B$2', '=$A1:$B2'])
    expect(cycleRefs('=A1+B2+C3', 3, 6, {})).toEqual({ text: '=A1+$B$2+C3', from: 3, to: 8 })
  })
  it('widens a selection whose ends fall inside references to the whole references it changed', () => {
    expect(cycleRefs('=A1+B2', 2, 5, {})).toEqual({ text: '=$A$1+$B$2', from: 1, to: 10 })
    expect(cycleRefs('=A1', 0, 2, {})).toEqual({ text: '=$A$1', from: 0, to: 5 })
  })
})
