// Excel's Formula AutoComplete: the name at the caret, the list under the cell's field and the formula bar's, its keys
// (arrows, Page keys, Home, End, Tab, Esc, Enter), a double click and a tap, and the source's names asked once per source.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import FunctionList from '../src/components/FunctionList.vue'
import { nameAt } from '../src/core/autocomplete'
import { active, editorKey, editorOpen, fakeSource, key, mountSheet, pressBar, settle, type } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

const FUNCTIONS = ['ЕСЛИ', 'СРЗНАЧ', 'СУММ', 'СУММЕСЛИ', 'СУММЕСЛИМН', 'СУММПРОИЗВ']
const book = (functions: string[] | null = FUNCTIONS): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 1, A2: 2 } }], ...(functions ? { functions } : {}) })

const field = (w: VueWrapper) => w.find('.xs-editor').element as HTMLTextAreaElement
const shown = () => [...document.querySelectorAll('[data-xs-function]')].map((li) => li.textContent!.trim())
const highlighted = () => document.querySelector('.xs-fl-on')?.textContent?.trim()
/** Open the cell's field with `=` and type the rest, as the browser sends it. */
async function typeFormula(w: VueWrapper, text: string): Promise<void> {
  await key(w, '=')
  await type(w, text)
  await settle()
}

describe('nameAt: the name typed at the caret', () => {
  it('after =, an operator, (, a separator or a space', () => {
    expect(nameAt('=СУ', 3)).toEqual({ from: 1, to: 3, typed: 'СУ' })
    expect(nameAt('=1+ЕС', 5)).toEqual({ from: 3, to: 5, typed: 'ЕС' })
    expect(nameAt('=СУММ(A1;СР', 11)).toEqual({ from: 9, to: 11, typed: 'СР' })
    expect(nameAt('=SUM(A1,AV', 10)).toEqual({ from: 8, to: 10, typed: 'AV' })
    expect(nameAt('=A1 * LOG1', 10)).toEqual({ from: 6, to: 10, typed: 'LOG1' })
    expect(nameAt('=(СТАНДОТКЛОН.', 14)).toEqual({ from: 2, to: 14, typed: 'СТАНДОТКЛОН.' })
  })

  it('takes the rest of the name after the caret into its place', () => {
    expect(nameAt('=СУММ(A1)', 3)).toEqual({ from: 1, to: 5, typed: 'СУ' })
  })

  it('none outside a formula, in a string, a quoted sheet, a table, after ! or $, or with nothing typed', () => {
    expect(nameAt('СУ', 2)).toBeNull()
    expect(nameAt('=', 1)).toBeNull()
    expect(nameAt('=СУММ(', 6)).toBeNull()
    expect(nameAt('="a СУ', 6)).toBeNull()
    expect(nameAt('="a""b"&СУ', 10)).toEqual({ from: 8, to: 10, typed: 'СУ' })
    expect(nameAt("='Q 1'!СУ", 9)).toBeNull()
    expect(nameAt("='Q (1'+СУ", 10)).toEqual({ from: 8, to: 10, typed: 'СУ' })
    expect(nameAt('=Продажи[Су', 11)).toBeNull()
    expect(nameAt('=Лист1!A', 8)).toBeNull()
    expect(nameAt('=$A', 3)).toBeNull()
  })
})

describe('the list under the cell', () => {
  it('shows the functions the typed name begins, in any case, the first highlighted; typing on narrows it, a character not of a name closes it', async () => {
    const w = await mountSheet(memorySource(book()))
    await typeFormula(w, '=су')
    expect(shown()).toEqual(['СУММ', 'СУММЕСЛИ', 'СУММЕСЛИМН', 'СУММПРОИЗВ'])
    expect(highlighted()).toBe('СУММ')
    expect(document.querySelector('[data-xs-functions]')!.getAttribute('aria-label')).toBe('Functions')
    await type(w, '=суммп')
    expect(shown()).toEqual(['СУММПРОИЗВ'])
    await type(w, '=суммпх')
    expect(shown()).toEqual([])
    await type(w, '=СУММ(')
    expect(shown()).toEqual([])
    await type(w, '=СУММ(A1;Е')
    expect(shown()).toEqual(['ЕСЛИ'])
    expect(editorOpen(w)).toBe(true)
  })

  it('the arrows, Page keys, Home and End move the highlight; Tab enters the name with "(" and closes the list', async () => {
    const w = await mountSheet(memorySource(book()))
    await typeFormula(w, '=СУ')
    await editorKey(w, 'ArrowDown')
    expect(highlighted()).toBe('СУММЕСЛИ')
    await editorKey(w, 'End')
    expect(highlighted()).toBe('СУММПРОИЗВ')
    await editorKey(w, 'ArrowDown')
    expect(highlighted()).toBe('СУММПРОИЗВ')
    await editorKey(w, 'Home')
    expect(highlighted()).toBe('СУММ')
    await editorKey(w, 'PageDown')
    expect(highlighted()).toBe('СУММПРОИЗВ')
    await editorKey(w, 'PageUp')
    await editorKey(w, 'ArrowDown')
    await editorKey(w, 'Tab')
    await settle()
    expect(shown()).toEqual([])
    expect(editorOpen(w)).toBe(true)
    expect(field(w).value).toBe('=СУММЕСЛИ(')
    expect(field(w).selectionStart).toBe(10)
    expect(active(w)).toBe('1:1')
  })

  it('Tab replaces the whole name and keeps a "(" that follows it', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, '=')
    const el = field(w)
    el.value = '=СУММ(A1)'
    el.setSelectionRange(3, 3)
    await w.find('.xs-editor').trigger('input')
    await settle()
    expect(shown()[0]).toBe('СУММ')
    await editorKey(w, 'ArrowDown')
    await editorKey(w, 'Tab')
    await settle()
    expect(el.value).toBe('=СУММЕСЛИ(A1)')
    expect(el.selectionStart).toBe(10)
  })

  it('Esc closes the list and keeps the field; the second Esc cancels the edit', async () => {
    const { source, edits } = recordedSource()
    const w = await mountSheet(source)
    await typeFormula(w, '=СУ')
    await editorKey(w, 'Escape')
    expect(shown()).toEqual([])
    expect(editorOpen(w)).toBe(true)
    expect(field(w).value).toBe('=СУ')
    await editorKey(w, 'Escape')
    await settle()
    expect(editorOpen(w)).toBe(false)
    expect(edits).toEqual([])
  })

  it('Enter with the list open commits the cell as typed and moves down, as Excel for Windows', async () => {
    const { source, edits } = recordedSource()
    const w = await mountSheet(source)
    await typeFormula(w, '=СУ')
    await editorKey(w, 'Enter')
    await settle()
    expect(edits).toEqual([[{ sheet: 'S', cell: 'A1', text: '=СУ' }]])
    expect(active(w)).toBe('2:1')
    expect(shown()).toEqual([])
  })

  it('Alt+Enter breaks the line and closes the list', async () => {
    const w = await mountSheet(memorySource(book()))
    await typeFormula(w, '=СУ')
    await editorKey(w, 'Enter', { altKey: true })
    await settle()
    expect(shown()).toEqual([])
    expect(field(w).value).toBe('=СУ\n')
  })

  it('a click highlights an item, a double click enters it', async () => {
    const w = await mountSheet(memorySource(book()))
    await typeFormula(w, '=СУ')
    await w.find('[data-xs-function="СУММПРОИЗВ"]').trigger('click')
    expect(highlighted()).toBe('СУММПРОИЗВ')
    expect(editorOpen(w)).toBe(true)
    await w.find('[data-xs-function="СУММЕСЛИ"]').trigger('dblclick')
    await settle()
    expect(field(w).value).toBe('=СУММЕСЛИ(')
    expect(shown()).toEqual([])
  })

  it('Left and Right move the caret and the list follows the name before it', async () => {
    const w = await mountSheet(memorySource(book()))
    await typeFormula(w, '=СУММ')
    expect(shown()).toHaveLength(4)
    // F2 to Edit mode, where the arrows move the caret: in Enter mode Left would enter `=С` (refAt's commit)
    await editorKey(w, 'F2')
    await new Promise((done) => setTimeout(done))
    await settle()
    expect(shown()).toHaveLength(4)
    // The browser moves the caret after the key; the list reads it once the key is done.
    field(w).setSelectionRange(2, 2)
    await editorKey(w, 'ArrowLeft')
    await new Promise((done) => setTimeout(done))
    await settle()
    expect(shown()).toEqual(['СРЗНАЧ', 'СУММ', 'СУММЕСЛИ', 'СУММЕСЛИМН', 'СУММПРОИЗВ'])
    field(w).setSelectionRange(1, 1)
    await editorKey(w, 'ArrowLeft')
    await new Promise((done) => setTimeout(done))
    await settle()
    expect(shown()).toEqual([])
  })

  it('no list for text that is not a formula, nor in a string', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, 'С')
    await type(w, 'СУ')
    await settle()
    expect(shown()).toEqual([])
    await editorKey(w, 'Escape')
    await typeFormula(w, '="СУ')
    expect(shown()).toEqual([])
  })
})

describe('the source’s names', () => {
  it('asked once per source, when a field opens, with a signal; a new source is asked again', async () => {
    const one = memorySource(book())
    const asked = vi.spyOn(one, 'functionNames')
    const w = await mountSheet(one)
    expect(asked).not.toHaveBeenCalled()
    await typeFormula(w, '=СУ')
    await editorKey(w, 'Escape')
    await editorKey(w, 'Escape')
    await typeFormula(w, '=ЕС')
    expect(shown()).toEqual(['ЕСЛИ'])
    expect(asked).toHaveBeenCalledTimes(1)
    expect(asked.mock.calls[0]![0]).toBeInstanceOf(AbortSignal)
    await editorKey(w, 'Escape')
    await editorKey(w, 'Escape')
    const two = memorySource(book(['SUM', 'SUMIF']))
    const asked2 = vi.spyOn(two, 'functionNames')
    await w.setProps({ source: two })
    await settle()
    await typeFormula(w, '=su')
    expect(asked2).toHaveBeenCalledTimes(1)
    expect(shown()).toEqual(['SUM', 'SUMIF'])
  })

  it('names that come after the typing open the list for what is typed then', async () => {
    let give: (names: string[]) => void = () => {}
    const fake = fakeSource([{ name: 'S', maxRow: 3, maxCol: 3 }])
    fake.source.functionNames = () => new Promise((resolve) => (give = resolve))
    const w = await mountSheet(fake.source)
    await typeFormula(w, '=СУ')
    expect(shown()).toEqual([])
    give(['СУММ', 'СРЗНАЧ'])
    await settle()
    expect(shown()).toEqual(['СУММ'])
  })

  it('a failed request is asked again with the next field', async () => {
    const fake = fakeSource([{ name: 'S', maxRow: 3, maxCol: 3 }])
    const names = vi.fn().mockRejectedValueOnce(new Error('offline')).mockResolvedValue(['СУММ'])
    fake.source.functionNames = names
    const w = await mountSheet(fake.source)
    await typeFormula(w, '=СУ')
    expect(shown()).toEqual([])
    await editorKey(w, 'Escape')
    await typeFormula(w, '=СУ')
    expect(names).toHaveBeenCalledTimes(2)
    expect(shown()).toEqual(['СУММ'])
  })

  it('without functionNames there is no list', async () => {
    const source = memorySource(book(null))
    expect(source.functionNames).toBeUndefined()
    const w = await mountSheet(source)
    await typeFormula(w, '=СУ')
    expect(shown()).toEqual([])
    expect(await memorySource(book()).functionNames!()).toEqual(FUNCTIONS)
  })
})

describe('the list under the formula bar', () => {
  it('shows under the bar’s field; Tab enters the name there', async () => {
    const w = await mountSheet(memorySource(book()))
    await pressBar(w, 0)
    const bar = w.find('.xs-fx-input')
    ;(bar.element as HTMLTextAreaElement).value = '=СР'
    await bar.trigger('input')
    await settle()
    expect(w.find('.xs-fx [data-xs-functions]').exists()).toBe(true)
    expect(shown()).toEqual(['СРЗНАЧ'])
    await bar.trigger('keydown', { key: 'Tab' })
    await settle()
    expect((bar.element as HTMLTextAreaElement).value).toBe('=СРЗНАЧ(')
    expect(shown()).toEqual([])
  })
})

describe('FunctionList', () => {
  const FIELD = { top: 100, left: 300, width: 80, height: 20 }
  const VIEW = { top: 0, left: 0, width: 360, height: 400 }

  it('under the field, eight items high, kept inside the view; over the field where the view has no room below', () => {
    const items = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    const w = mount(FunctionList, { props: { rect: FIELD, items, view: VIEW } })
    const style = (w.element as HTMLElement).style
    expect([style.top, style.left, style.width, style.maxHeight]).toEqual(['120px', '140px', '220px', `${8 * 18 + 2}px`])
    const low = mount(FunctionList, { props: { rect: { ...FIELD, top: 350 }, items, view: VIEW } })
    expect((low.element as HTMLElement).style.top).toBe(`${350 - 8 * 18 - 2}px`)
  })

  it('under a finger the items are 44 px and a tap enters one; with the mouse a click highlights', async () => {
    const w = mount(FunctionList, { props: { rect: FIELD, items: ['СУММ', 'СУММЕСЛИ'], touch: true, messages: { functionList: 'Функции' } } })
    expect(w.classes()).toContain('xs-fl-touch')
    expect((w.element as HTMLElement).style.maxHeight).toBe(`${2 * 44 + 2}px`)
    expect(w.attributes('aria-label')).toBe('Функции')
    await w.find('[data-xs-function="СУММЕСЛИ"]').trigger('click')
    expect(w.emitted('pick')).toEqual([[1]])
    const mouse = mount(FunctionList, { props: { rect: FIELD, items: ['СУММ', 'СУММЕСЛИ'] } })
    await mouse.find('[data-xs-function="СУММЕСЛИ"]').trigger('click')
    expect(mouse.emitted('highlight')).toEqual([[1]])
    expect(mouse.emitted('pick')).toBeUndefined()
  })

  it('a press on the list keeps the field’s focus', async () => {
    const w = mount(FunctionList, { props: { rect: FIELD, items: ['СУММ'] }, attachTo: document.body })
    const down = new MouseEvent('mousedown', { bubbles: true, cancelable: true })
    w.find('li').element.dispatchEvent(down)
    expect(down.defaultPrevented).toBe(true)
  })
})

/** A memory source of the book that records the edits it gets. */
function recordedSource() {
  const source = memorySource(book())
  const edits: unknown[] = []
  const edit = source.edit!.bind(source)
  source.edit = (entries, opts) => {
    edits.push(entries)
    return edit(entries, opts)
  }
  return { source, edits }
}
