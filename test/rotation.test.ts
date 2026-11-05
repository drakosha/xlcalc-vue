// Excel's text orientation (Format Cells > Alignment > Orientation): rotated and vertical text drawn in its cell, and the
// dialog's Degrees and Vertical Text, sent as the alignment's `rotation` only when changed.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { FormatDialog, memorySource, ruMessages, type SheetViewState } from '../src'
import { rotatedBox } from '../src/core/rotation'
import { draftErrors, patchesOf, startOf, type FormatDraft } from '../src/format-cells'
import { mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const at = (more: Partial<SheetViewState> = {}): SheetViewState => ({ sheet: 'S', ref: 'B2', ...more })
function ok(state: SheetViewState, change: (d: FormatDraft) => void) {
  const start = startOf(state)
  const draft = structuredClone(start.draft)
  change(draft)
  return patchesOf(start, draft)
}

describe("the dialog's Orientation", () => {
  it("opens with the cell's degrees, or with Vertical Text pressed", () => {
    expect(startOf(at()).draft).toMatchObject({ degrees: '0', verticalText: false })
    expect(startOf(at({ style: { alignment: { rotation: 45 } } })).draft).toMatchObject({ degrees: '45', verticalText: false })
    expect(startOf(at({ style: { alignment: { rotation: -90 } } })).draft).toMatchObject({ degrees: '-90', verticalText: false })
    expect(startOf(at({ style: { alignment: { rotation: 'vertical' } } })).draft).toMatchObject({ degrees: '0', verticalText: true })
  })

  it('sends the rotation only when it changed; 0 removes it, as Excel stores no default', () => {
    expect(ok(at(), (d) => (d.degrees = '30'))).toEqual([{ style: { alignment: { rotation: 30 } } }])
    expect(ok(at(), (d) => (d.verticalText = true))).toEqual([{ style: { alignment: { rotation: 'vertical' } } }])
    expect(ok(at({ style: { alignment: { rotation: 45 } } }), (d) => (d.degrees = '0'))).toEqual([{ style: { alignment: { rotation: null } } }])
    expect(ok(at({ style: { alignment: { rotation: 'vertical' } } }), (d) => (d.verticalText = false))).toEqual([{ style: { alignment: { rotation: null } } }])
    expect(ok(at({ style: { alignment: { rotation: 45 } } }), (d) => (d.degrees = ' 45 '))).toEqual([])
  })

  it("blocks OK outside Excel's −90 to 90 whole degrees", () => {
    const draft = startOf(at()).draft
    for (const bad of ['91', '-91', '1.5', 'x', '']) expect(draftErrors({ ...draft, degrees: bad }, draft.numberCode).degrees, bad).toBe(true)
    for (const good of ['90', '-90', '0', ' 15 ']) expect(draftErrors({ ...draft, degrees: good }, draft.numberCode).degrees, good).toBeUndefined()
    // Vertical Text leaves the degrees out of it
    expect(draftErrors({ ...draft, degrees: 'x', verticalText: true }, draft.numberCode).degrees).toBeUndefined()
  })
})

describe("FormatDialog's Orientation", () => {
  async function open(state: SheetViewState, messages?: typeof ruMessages): Promise<VueWrapper> {
    const w = mount(FormatDialog, { props: { open: true, state, messages }, attachTo: document.body })
    await settle()
    await w.find('[data-fmt-tab="alignment"]').trigger('click')
    return w
  }

  it('has Degrees and Vertical Text on the Alignment tab: OK sends the rotation, Vertical Text sets the degrees aside', async () => {
    const w = await open(at({ style: { alignment: { rotation: 45 } } }))
    const degrees = w.find('[data-fmt-degrees]')
    expect((degrees.element as HTMLInputElement).value).toBe('45')
    await degrees.setValue('-30')
    await w.find('[data-fmt-ok]').trigger('submit')
    expect(w.emitted('apply')).toEqual([[[{ style: { alignment: { rotation: -30 } } }]]])

    const v = await open(at())
    await v.find('[data-fmt-vertical-text]').setValue(true)
    expect(v.find('[data-fmt-degrees]').attributes('disabled')).toBeDefined()
    await v.find('[data-fmt-ok]').trigger('submit')
    expect(v.emitted('apply')).toEqual([[[{ style: { alignment: { rotation: 'vertical' } } }]]])
  })

  it('blocks OK with degrees out of range and says why, in the interface’s language', async () => {
    const w = await open(at(), ruMessages)
    await w.find('[data-fmt-degrees]').setValue('120')
    expect(w.find('[data-fmt-ok]').attributes('disabled')).toBeDefined()
    expect(w.text()).toContain('от -90 до 90')
    expect(w.text()).toContain('Ориентация')
  })
})

describe('rotated text in its cell', () => {
  it('takes the box of the rotated line: tall and narrow at 90°, as wide as long at 0°, between them at 45°', () => {
    const font = { family: 'Arial', px: 10, bold: false, italic: false }
    const up = rotatedBox('abcdefghij', font, 90)
    expect(up.width).toBe(13)
    expect(up.height).toBeGreaterThan(40)
    const level = rotatedBox('abcdefghij', font, 0)
    expect(level.height).toBe(13)
    const mid = rotatedBox('abcdefghij', font, -45)
    expect(mid.width).toBe(mid.height)
    expect(mid.width).toBeGreaterThan(up.width)
    expect(mid.angle).toBe(-45)
  })

  async function sheet(rotation: number | 'vertical' | undefined): Promise<VueWrapper> {
    const src = memorySource({ styles: [{ alignment: { rotation } }], sheets: [{ name: 'S', cells: { A1: { value: 'Rotated text', style: 0 } } }] })
    return mountSheet(src)
  }
  const cell = (w: VueWrapper) => w.find('[data-cell="1:1"]')

  it('draws text at an angle rotated about its box, counterclockwise for positive degrees, without overflowing', async () => {
    const w = await sheet(60)
    const text = cell(w).find('.xs-rot-text')
    expect(text.text()).toBe('Rotated text')
    expect(text.attributes('style')).toContain('rotate(-60deg)')
    expect(cell(w).classes()).not.toContain('xs-spill')
    const box = cell(w).find('.xs-rot-box').element as HTMLElement
    expect(parseFloat(box.style.height)).toBeGreaterThan(parseFloat(box.style.width))
  })

  it('stacks vertical text letter above letter', async () => {
    const w = await sheet('vertical')
    expect(cell(w).find('.xs-vert').text()).toBe('Rotated text')
    expect(cell(w).find('.xs-rot-text').exists()).toBe(false)
  })

  it('draws level text as before', async () => {
    const w = await sheet(0)
    expect(cell(w).find('.xs-rot').exists()).toBe(false)
    expect(cell(w).classes()).toContain('xs-spill')
  })
})
