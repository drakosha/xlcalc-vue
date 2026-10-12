// @vitest-environment node
// The Format Cells dialog's rules: its fields as the active cell opens them, the Border tab's buttons, the limits OK
// checks and the patches it sends, only for what changed, the borders on their ranges in the same call.
import { describe, expect, it } from 'vitest'
import { numberFormatChoices } from 'xlcalc/text'
import type { SheetViewState } from '../src'
import { draftErrors, patchesOf, pickNumber, pressEdge, pressPreset, shownEdge, startOf, typeNumber, type FormatDraft } from '../src/format-cells'

/** Excel's list as xlcalc gives it in en-US. */
const item = (key: string) => numberFormatChoices().find((c) => c.key === key)!
const THIN = { style: 'thin' } as const
const RED = { style: 'medium', color: '#C00000' } as const
const at = (ref: string, more: Partial<SheetViewState> = {}): SheetViewState => ({ sheet: 'S', ref, ...more })
/** The dialog opened on `state`, its fields changed by `change`, and OK's patches. */
function ok(state: SheetViewState, change: (d: FormatDraft, s: ReturnType<typeof startOf>) => void) {
  const start = startOf(state)
  const draft = structuredClone(start.draft)
  change(draft, start)
  return patchesOf(start, draft)
}

describe('the fields as the active cell opens them', () => {
  it('takes the font, the alignment, the fill, the number format among those offered and the cell’s own sides', () => {
    const s = startOf(
      at('B2', {
        numberFormat: '0%',
        style: {
          font: { name: 'Arial', size: 14, bold: true, italic: true, underline: 'double', strike: true, color: '#ff0000' },
          alignment: { horizontal: 'center', vertical: 'top', wrap: true, indent: 2 },
          fill: { color: '#ffff00' },
          border: { top: THIN },
        },
      }),
    )
    expect(s.draft).toMatchObject({ numberCode: '0%', numberFormat: null, horizontal: 'center', vertical: 'top', wrap: true, indent: '2', fontName: 'Arial', fontSize: '14', fontStyle: 'boldItalic', underline: 'double', strike: true, fontColour: '#FF0000', fill: '#FFFF00' })
    expect(s.single).toBe(true)
    expect(s.edges).toEqual({ top: THIN, bottom: null, left: null, right: null })
  })

  it('opens a cell without a look at Excel’s defaults, a code of its own as Custom, a range’s edges unknown', () => {
    expect(startOf(at('A1')).draft).toMatchObject({ numberCode: 'General', numberFormat: null, horizontal: 'general', vertical: 'bottom', wrap: false, indent: '0', fontName: '', fontSize: '', fontStyle: 'regular', underline: 'none', fontColour: null, fill: null })
    expect(startOf(at('A1', { numberFormat: '0.000' })).draft).toMatchObject({ numberCode: '0.000' })
    const range = startOf(at('B2:D4', { style: { border: { top: THIN } } }))
    expect([range.single, range.edges]).toEqual([false, {}])
  })
})

describe('OK’s patches', () => {
  it('nothing changed, nothing sent', () => {
    expect(ok(at('B2', { style: { font: { bold: true } }, numberFormat: '0.000' }), () => {})).toEqual([])
  })

  it('the font as set, over the selection; a size to Excel’s half points', () => {
    const patches = ok(at('B2:C3'), (d) => {
      Object.assign(d, { fontName: 'Arial', fontSize: '10,3', fontStyle: 'boldItalic', underline: 'single', strike: true, fontColour: '#C00000' })
    })
    expect(patches).toEqual([{ style: { font: { name: 'Arial', size: 10.5, bold: true, italic: true, underline: 'single', strike: true, color: '#C00000' } } }])
  })

  it('a field back at Excel’s default takes the cell’s own off', () => {
    const state = at('B2', { style: { font: { bold: true, color: '#FF0000', size: 14 }, alignment: { horizontal: 'center', vertical: 'top', wrap: true, indent: 1 }, fill: { color: '#FFFF00' } } })
    const patches = ok(state, (d) => {
      Object.assign(d, { fontStyle: 'regular', fontColour: null, fontSize: '', horizontal: 'general', vertical: 'bottom', wrap: false, indent: '0', fill: null })
    })
    expect(patches).toEqual([{ style: { font: { size: null, bold: null, color: null }, alignment: { horizontal: null, vertical: null, wrap: null, indent: null }, fill: null } }])
  })

  it('the alignment and the fill as set', () => {
    expect(ok(at('A1'), (d) => Object.assign(d, { horizontal: 'right', vertical: 'center', wrap: true, indent: '3', fill: '#92D050' }))).toEqual([
      { style: { alignment: { horizontal: 'right', vertical: 'center', wrap: true, indent: 3 }, fill: { color: '#92D050' } } },
    ])
  })

  it('an item of Excel’s list by what it sets, a code typed as written; the same code sends nothing', () => {
    expect(ok(at('A1:A3'), (d) => pickNumber(d, item('percentage')))).toEqual([{ numberFormat: 10 }])
    expect(ok(at('A1'), (d) => typeNumber(d, ' 0.000 '))).toEqual([{ numberFormat: '0.000' }])
    expect(ok(at('A1', { numberFormat: '0.000' }), (d) => typeNumber(d, '0.000'))).toEqual([])
    expect(ok(at('A1', { numberFormat: '0.00%' }), (d) => pickNumber(d, item('percentage')))).toEqual([])
    expect(ok(at('A1', { numberFormat: '0.000' }), (d) => pickNumber(d, item('general')))).toEqual([{ numberFormat: 0 }])
  })

  it('an item puts its code in the field, to be changed from there as Excel’s Custom', () => {
    const patches = ok(at('A1'), (d) => {
      pickNumber(d, item('percentage'))
      expect(d.numberCode).toBe('0.00%')
      typeNumber(d, '0.0%')
    })
    expect(patches).toEqual([{ numberFormat: '0.0%' }])
  })
})

describe('the Border tab', () => {
  it('an edge’s button draws the line, or takes it off when the edge has that line; one cell’s sides on the cell', () => {
    const patches = ok(at('B2', { style: { border: { top: THIN } } }), (d, s) => {
      pressEdge(s, d, 'top')
      pressEdge(s, d, 'bottom')
      expect([shownEdge(s, d, 'top'), shownEdge(s, d, 'bottom')]).toEqual([null, THIN])
    })
    expect(patches).toEqual([
      { range: 'B2', style: { border: { top: null } } },
      { range: 'B2', style: { border: { bottom: THIN } } },
    ])
  })

  it('Outline and Inside with one line make one patch over every cell; None takes every line off at once', () => {
    expect(ok(at('B2:D4'), (d, s) => {
      d.line = { ...RED }
      pressPreset(s, d, 'outline')
      pressPreset(s, d, 'inside')
    })).toEqual([{ range: 'B2:D4', style: { border: { top: RED, right: RED, bottom: RED, left: RED } } }])
    expect(ok(at('B2:D4'), (d, s) => pressPreset(s, d, 'none'))).toEqual([{ range: 'B2:D4', style: { border: null } }])
  })

  it('the outline as four patches, a line inside on both sides of the cells it lies between, in one call', () => {
    expect(ok(at('B2:D4'), (d, s) => pressPreset(s, d, 'outline'))).toEqual([
      { range: 'B2:D2', style: { border: { top: THIN } } },
      { range: 'B4:D4', style: { border: { bottom: THIN } } },
      { range: 'B2:B4', style: { border: { left: THIN } } },
      { range: 'D2:D4', style: { border: { right: THIN } } },
    ])
    expect(ok(at('B2:D4'), (d, s) => pressEdge(s, d, 'insideH'))).toEqual([
      { range: 'B2:D3', style: { border: { bottom: THIN } } },
      { range: 'B3:D4', style: { border: { top: THIN } } },
    ])
  })

  it('a range’s edge pressed twice is taken off, as Excel’s grey button; one cell has no inside lines', () => {
    expect(ok(at('B2:C3'), (d, s) => {
      pressEdge(s, d, 'insideV')
      pressEdge(s, d, 'insideV')
    })).toEqual([
      { range: 'B2:B3', style: { border: { right: null } } },
      { range: 'C2:C3', style: { border: { left: null } } },
    ])
    expect(ok(at('B2'), (d, s) => pressPreset(s, d, 'inside'))).toEqual([])
    // a cell's side known to have no line, drawn and taken off again, is as it was
    expect(ok(at('B2'), (d, s) => {
      pressEdge(s, d, 'left')
      pressEdge(s, d, 'left')
    })).toEqual([])
  })

  it('the font and the borders together in one call', () => {
    expect(ok(at('A1:B1'), (d, s) => {
      d.fontStyle = 'bold'
      pressEdge(s, d, 'bottom')
    })).toEqual([{ style: { font: { bold: true } } }, { range: 'A1:B1', style: { border: { bottom: THIN } } }])
  })
})

describe('the limits OK checks', () => {
  const errors = (d: Partial<FormatDraft>) => draftErrors({ ...startOf(at('A1')).draft, ...d })
  it('a font of 1 to 409 points, an indent of 0 to 250, a code of the cell’s own not empty', () => {
    expect(errors({ fontSize: '10,5', indent: '250' })).toEqual({})
    expect(errors({ fontSize: '0', indent: '251' })).toEqual({ fontSize: true, indent: true })
    expect(errors({ fontSize: '409.5', indent: '-1' })).toEqual({ fontSize: true, indent: true })
    expect(errors({ fontSize: 'big', indent: '1.5' })).toEqual({ fontSize: true, indent: true })
    expect(errors({ numberCode: '  ' })).toEqual({ numberCode: true })
  })
})
