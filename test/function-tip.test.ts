// Excel's argument tooltip while a formula is typed: the call at the caret (the innermost), the argument at the caret in bold
// (a repeating group counted up), below the cell editor and the formula bar's, never together with the function list; the
// source's signatures requested by name, once, with a stale request aborted.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { functionSignature } from 'xlcalc'
import { ruRUFunctionText } from 'xlcalc/function-text/ru-RU'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { memorySource, type MemoryBook, type SheetFunctionSignature } from '../src'
import FunctionTip from '../src/components/FunctionTip.vue'
import { tipParts } from '../src/core/function-tip'
import { editorKey, editorOpen, key, mountSheet, pressBar, settle, type } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

const book = (functions?: string[]): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 1 } }], ...(functions ? { functions } : {}) })
/** An in-memory source with the library's signatures in a locale. */
function signed(locale?: typeof ruRU, functions?: string[]) {
  const source = memorySource(book(functions))
  const asked = vi.fn(async (name: string, _signal?: AbortSignal) => functionSignature(name, locale && { ...locale, ...ruRUFunctionText }))
  source.functionSignature = asked
  return { source, asked }
}

const tip = () => document.querySelector('[data-xs-tip]')?.textContent?.trim() ?? null
const current = () => document.querySelector('[data-xs-tip] .xs-ft-on')?.textContent?.trim() ?? null
const field = (w: VueWrapper) => w.find('.xs-editor').element as HTMLTextAreaElement
async function typeFormula(w: VueWrapper, text: string): Promise<void> {
  await key(w, '=')
  await type(w, text)
  await settle()
}
/** Puts the caret at `at` in the editor as a click does. */
async function caretTo(w: VueWrapper, at: number): Promise<void> {
  field(w).setSelectionRange(at, at)
  await w.find('.xs-editor').trigger('click')
  await new Promise((done) => setTimeout(done))
  await settle()
}

describe('tipParts: the tooltip in pieces with the current argument', () => {
  const sum = functionSignature('SUM')!
  const sumifs = functionSignature('SUMIFS')!
  const on = (parts: ReturnType<typeof tipParts>) => parts.filter((p) => p.current).map((p) => p.text)
  const text = (parts: ReturnType<typeof tipParts>) => parts.map((p) => p.text).join('')

  it('an argument within the list', () => {
    expect(text(tipParts(sum, 0))).toBe('SUM(number1, [number2], ...)')
    expect(on(tipParts(sum, 0))).toEqual(['number1'])
    expect(on(tipParts(sum, 1))).toEqual(['number2'])
    expect(on(tipParts(sumifs, 2))).toEqual(['criteria1'])
  })

  it('beyond the list the repeating group continues, with its numbers counted up', () => {
    expect(text(tipParts(sum, 2))).toBe('SUM(number1, [number2], [number3], ...)')
    expect(on(tipParts(sum, 2))).toEqual(['number3'])
    expect(on(tipParts(sum, 4))).toEqual(['number5'])
    expect(text(tipParts(sumifs, 5))).toBe('SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], [criteria_range3, criteria3], ...)')
    expect(on(tipParts(sumifs, 5))).toEqual(['criteria_range3'])
    expect(on(tipParts(sumifs, 6))).toEqual(['criteria3'])
    expect(text(tipParts(functionSignature('СУММ', { ...ruRU, ...ruRUFunctionText })!, 2))).toBe('СУММ(число1;[число2];[число3];...)')
  })

  it('beyond the last argument of a function that takes no more: nothing in bold', () => {
    const round = functionSignature('ROUND')!
    expect(text(tipParts(round, 2))).toBe(round.text)
    expect(on(tipParts(round, 2))).toEqual([])
  })

  it("a host's own signature, with names its text does not contain, is displayed as its text", () => {
    const own: SheetFunctionSignature = { name: 'F', fn: 'F', args: [{ name: 'x', optional: false }], repeat: 0, text: 'F(value)', category: 'math', categoryName: 'Math & Trig', description: 'The host’s own.' }
    expect(tipParts(own, 0)).toEqual([{ text: 'F(value)' }])
  })
})

describe('the tooltip below the cell', () => {
  it('in `=SUM(` the first argument, and the next one after a separator', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(')
    expect(tip()).toBe('SUM(number1, [number2], ...)')
    expect(current()).toBe('number1')
    await type(w, '=SUM(1, 2, ')
    await settle()
    expect(tip()).toBe('SUM(number1, [number2], [number3], ...)')
    expect(current()).toBe('number3')
  })

  it('SUMIFS: the second group of criteria, then the third, counted up', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await typeFormula(w, '=SUMIFS(A1:A3, B1:B3, "x", ')
    expect(current()).toBe('criteria_range2')
    await type(w, '=SUMIFS(A1:A3, B1:B3, "x", C1:C3, "y", ')
    await settle()
    expect(current()).toBe('criteria_range3')
    expect(tip()).toContain('[criteria_range3, criteria3]')
  })

  it('ru-RU: `=СУММ(1;` is at the second argument; the source is queried by the English name', async () => {
    const { source, asked } = signed(ruRU)
    const w = await mountSheet(source, { locale: 'ru-RU' })
    await settle()
    await typeFormula(w, '=СУММ(1;')
    expect(tip()).toBe('СУММ(число1;[число2];...)')
    expect(current()).toBe('число2')
    // the source is queried by the English name, once for every spelling of it
    await type(w, '=СУММ(1;сумм(')
    await settle()
    expect(current()).toBe('число1')
    expect(asked.mock.calls.map((c) => c[0])).toEqual(['SUM'])
  })

  it('a call inside a call: the inner one', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(1, ROUND(')
    expect(tip()).toMatch(/^ROUND\(/)
    expect(current()).toBe('number')
    await type(w, '=SUM(1, ROUND(2, 1), ')
    await settle()
    expect(tip()).toMatch(/^SUM\(/)
    expect(current()).toBe('number3')
  })

  it('none in a string, after the closing parenthesis, outside a formula, or for an unknown name', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await typeFormula(w, '="SUM(')
    expect(tip()).toBeNull()
    await type(w, '=SUM(1)')
    await settle()
    expect(tip()).toBeNull()
    await type(w, '=SUM(1)+')
    await settle()
    expect(tip()).toBeNull()
    await type(w, '=NOSUCH(')
    await settle()
    expect(tip()).toBeNull()
    await editorKey(w, 'Escape')
    await key(w, 'S')
    await type(w, 'SUM(')
    await settle()
    expect(tip()).toBeNull()
  })

  it('follows the caret: it appears back inside the parentheses and closes outside them', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(1)')
    expect(tip()).toBeNull()
    await caretTo(w, 6)
    expect(current()).toBe('number1')
    await caretTo(w, 7)
    expect(tip()).toBeNull()
  })

  it('without functionSignature there is none', async () => {
    const source = memorySource(book())
    expect(source.functionSignature).toBeUndefined()
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(')
    expect(tip()).toBeNull()
  })

  it('never together with the function list: the list first, the tooltip once it closes', async () => {
    const { source } = signed(undefined, ['ROUND', 'SUM', 'SUMIF'])
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(1, RO')
    expect(document.querySelector('[data-xs-functions]')).not.toBeNull()
    expect(tip()).toBeNull()
    await editorKey(w, 'Escape')
    await settle()
    expect(document.querySelector('[data-xs-functions]')).toBeNull()
    expect(current()).toBe('number2')
  })

  it('Esc cancels the edit along with the tooltip; Enter commits and closes it', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(')
    await editorKey(w, 'Escape')
    await settle()
    expect(editorOpen(w)).toBe(false)
    expect(tip()).toBeNull()
    await typeFormula(w, '=SUM(1')
    expect(tip()).not.toBeNull()
    await editorKey(w, 'Enter')
    await settle()
    expect(tip()).toBeNull()
  })
})

describe("the source's signatures", () => {
  it('requested by name once, not on every keystroke, with a signal; any letter case of a name counts as one', async () => {
    const { source, asked } = signed()
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(')
    await type(w, '=SUM(1')
    await type(w, '=SUM(1, ')
    await type(w, '=SUM(1, 2')
    await settle()
    expect(current()).toBe('number2')
    await type(w, '=SUM(1, 2) + sum(')
    await settle()
    expect(current()).toBe('number1')
    expect(asked).toHaveBeenCalledTimes(1)
    expect(asked.mock.calls[0]![0]).toBe('SUM')
    expect(asked.mock.calls[0]![1]).toBeInstanceOf(AbortSignal)
  })

  it('a response for a call the caret has left is aborted and not shown', async () => {
    const { source } = signed()
    const signals: AbortSignal[] = []
    const gives: ((s: SheetFunctionSignature | null) => void)[] = []
    source.functionSignature = (_name, signal) => {
      signals.push(signal!)
      return new Promise((resolve) => gives.push(resolve))
    }
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(')
    await type(w, '=SUM(1, ROUND(')
    await settle()
    expect(signals).toHaveLength(2)
    expect(signals[0]!.aborted).toBe(true)
    gives[0]!(functionSignature('SUM'))
    await settle()
    expect(tip()).toBeNull()
    gives[1]!(functionSignature('ROUND'))
    await settle()
    expect(tip()).toMatch(/^ROUND\(/)
  })

  it('a failed request shows nothing and is retried later', async () => {
    const { source } = signed()
    const asked = vi.fn().mockRejectedValueOnce(new Error('offline')).mockImplementation(async (name: string) => functionSignature(name))
    source.functionSignature = asked
    const w = await mountSheet(source)
    await typeFormula(w, '=SUM(')
    expect(tip()).toBeNull()
    await editorKey(w, 'Escape')
    await typeFormula(w, '=SUM(')
    expect(asked).toHaveBeenCalledTimes(2)
    expect(current()).toBe('number1')
  })
})

describe('the tooltip below the formula bar', () => {
  it("appears below the formula bar's input", async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await pressBar(w, 0)
    await settle()
    const bar = w.find('.xs-fx-input')
    ;(bar.element as HTMLTextAreaElement).value = '=SUMIFS(A1:A3, B1:B3, '
    await bar.trigger('input')
    await settle()
    expect(w.find('.xs-fx [data-xs-tip]').exists()).toBe(true)
    expect(current()).toBe('criteria1')
  })
})

describe('FunctionTip', () => {
  const FIELD = { top: 100, left: 300, width: 80, height: 20 }
  const VIEW = { top: 0, left: 0, width: 360, height: 400 }
  const parts = tipParts(functionSignature('СУММ', { ...ruRU, ...ruRUFunctionText })!, 1)

  it('below the editor, kept inside the viewport by its measured width; above the editor where the viewport has no room below', async () => {
    vi.spyOn(HTMLElement.prototype, 'offsetWidth', 'get').mockReturnValue(300)
    vi.spyOn(HTMLElement.prototype, 'offsetHeight', 'get').mockReturnValue(36)
    const w = mount(FunctionTip, { props: { rect: FIELD, parts, view: VIEW } })
    await settle()
    const style = (w.element as HTMLElement).style
    expect([style.top, style.left, style.maxWidth]).toEqual(['120px', '60px', '360px'])
    const low = mount(FunctionTip, { props: { rect: { ...FIELD, top: 380 }, parts, view: VIEW } })
    await settle()
    expect((low.element as HTMLElement).style.top).toBe(`${380 - 36}px`)
  })

  it('the current argument in bold; a narrow viewport may break the line after a separator or "(", not inside a name', () => {
    const w = mount(FunctionTip, { props: { rect: FIELD, parts } })
    expect(w.find('.xs-ft-on').text()).toBe('число2')
    expect(w.html().replace(/<!--.*?-->|\s/g, '')).toContain('СУММ(<wbr>число1;<wbr>[<b')
    expect(w.text()).toBe('СУММ(число1;[число2];...)')
  })

  it("a press on the tooltip keeps the editor's focus", () => {
    const w = mount(FunctionTip, { props: { rect: FIELD, parts }, attachTo: document.body })
    const down = new MouseEvent('mousedown', { bubbles: true, cancelable: true })
    w.element.dispatchEvent(down)
    expect(down.defaultPrevented).toBe(true)
  })
})
