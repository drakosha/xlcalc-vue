// Excel's Spelling (F7, checkSpelling) over memorySource with a word list: the dialog from the active cell with the cell of
// each word selected, Ignore Once, Ignore All, Add to Dictionary, Change and Change All through `edit`, within a selection,
// and Excel's message once the check comes back to where it started.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook, type SheetViewApi } from '../src'
import type { SpellChecker } from '../src/contract/workbook'
import { active, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const WORDS = new Set(['good', 'bad', 'word', 'here', 'again', 'the', 'end', 'red'])
const SUGGEST: Record<string, string[]> = { bda: ['bad', 'bead'], wrod: ['word'], teh: ['the'] }
function checker(): SpellChecker & { add: ReturnType<typeof vi.fn> } {
  const own = new Set<string>()
  return {
    check: (w) => WORDS.has(w.toLowerCase()) || own.has(w.toLowerCase()),
    suggest: (w) => SUGGEST[w.toLowerCase()] ?? [],
    add: vi.fn((w: string) => void own.add(w.toLowerCase())),
  }
}
const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'good bda', B1: 'wrod here', A2: 'teh end', B3: 'bda again', A4: 'zzq' } }] })

const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const dialog = (w: VueWrapper) => w.find('[data-xs-spelling]')
const word = (w: VueWrapper) => w.find('.xs-spell-word').text()
const button = async (w: VueWrapper, name: string) => {
  await w.find(`[data-xs-spell="${name}"]`).trigger('click')
  await settle(10)
}
const done = (w: VueWrapper) => w.find('[data-xs-spell-done]').exists()

describe('Spelling', () => {
  it('F7 starts at the active cell and selects the cell of each word; Change writes the cell, Ignore Once goes on, and the end is Excel’s message', async () => {
    const source = memorySource(book(), { spelling: checker() })
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await key(w, 'F7')
    await settle(10)
    expect(dialog(w).exists()).toBe(true)
    expect(w.find('[data-xs-spell-text]').text()).toBe('good bda')
    expect(word(w)).toBe('bda')
    expect(w.findAll('[data-xs-spell-option]').map((o) => o.text())).toEqual(['bad', 'bead'])
    expect(active(w)).toBe('1:1')
    await button(w, 'change')
    expect(edit).toHaveBeenLastCalledWith([{ sheet: 'S', cell: 'A1', text: 'good bad' }])
    expect(word(w)).toBe('wrod')
    expect(active(w)).toBe('1:2')
    await button(w, 'ignoreOnce')
    expect(word(w)).toBe('teh')
    await button(w, 'ignoreOnce')
    expect(word(w)).toBe('bda')
    expect(active(w)).toBe('3:2')
    await button(w, 'ignoreOnce')
    // no suggestions for zzq: Change is disabled
    expect(word(w)).toBe('zzq')
    expect(w.find('.xs-spell-none').text()).toBe('(No Suggestions)')
    expect(w.find('[data-xs-spell="change"]').attributes('disabled')).toBeDefined()
    await button(w, 'ignoreOnce')
    expect(dialog(w).exists()).toBe(false)
    expect(done(w)).toBe(true)
    expect(w.find('[data-xs-spell-done]').text()).toContain('Spell check complete')
    await w.find('[data-xs-dialog-ok]').trigger('click')
    await settle()
    expect(done(w)).toBe(false)
  })

  it('Change All replaces the word in the cells checked later as they are found; Ignore All skips it for the session', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 'bda wrod', A2: 'wrod', A3: 'bda bda', A4: 'BDA' } }] }, { spelling: checker() })
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    expect(api(w).checkSpelling()).toBe(true)
    await settle(10)
    expect(word(w)).toBe('bda')
    await button(w, 'changeAll')
    // A1 itself, then wrod; A3's two words are changed without stopping when the check reaches them
    expect(edit.mock.calls.map((c) => c[0][0]!.text)).toEqual(['bad wrod'])
    expect(word(w)).toBe('wrod')
    await button(w, 'ignoreAll')
    expect(edit.mock.calls.map((c) => c[0][0]!)).toEqual([
      { sheet: 'S', cell: 'A1', text: 'bad wrod' },
      { sheet: 'S', cell: 'A3', text: 'bad bda' },
      { sheet: 'S', cell: 'A3', text: 'bad bad' },
    ])
    // A2's wrod is ignored, and BDA in uppercase is not checked, as Excel's default
    expect(done(w)).toBe(true)
  })

  it('Add to Dictionary teaches the word to the source and goes on', async () => {
    const c = checker()
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 'zzq', A2: 'zzq too' } }] }, { spelling: c }))
    api(w).checkSpelling()
    await settle(10)
    await button(w, 'add')
    expect(c.add).toHaveBeenCalledWith('zzq')
    expect(word(w)).toBe('too')
  })

  it('starts at the active cell and wraps around to the cells before it; within a selection of several cells, only those', async () => {
    const w = await mountSheet(memorySource(book(), { spelling: checker() }))
    await goTo(w, 'A2')
    api(w).checkSpelling()
    await settle(10)
    expect(word(w)).toBe('teh')
    await button(w, 'ignoreOnce')
    expect(word(w)).toBe('bda')
    await button(w, 'ignoreOnce')
    expect(word(w)).toBe('zzq')
    await button(w, 'ignoreOnce')
    expect(word(w)).toBe('bda')
    expect(active(w)).toBe('1:1')
    await button(w, 'ignoreOnce')
    expect(word(w)).toBe('wrod')
    await button(w, 'ignoreOnce')
    expect(done(w)).toBe(true)
    await w.find('[data-xs-dialog-ok]').trigger('click')
    await goTo(w, 'A2:B3')
    api(w).checkSpelling()
    await settle(10)
    expect(word(w)).toBe('teh')
    await button(w, 'ignoreOnce')
    expect(word(w)).toBe('bda')
    await button(w, 'ignoreOnce')
    expect(done(w)).toBe(true)
  })

  it('a sheet without misspelt words gets the message at once; a read-only sheet cannot Change; no `spelling`, no Spelling', async () => {
    const clean = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 'good word' } }] }, { spelling: checker() }))
    api(clean).checkSpelling()
    await settle(10)
    expect(dialog(clean).exists()).toBe(false)
    expect(done(clean)).toBe(true)
    const ro = await mountSheet(memorySource(book(), { spelling: checker() }), { readOnly: true })
    api(ro).checkSpelling()
    await settle(10)
    expect(ro.find('[data-xs-spell="change"]').attributes('disabled')).toBeDefined()
    expect(ro.find('[data-xs-spell="ignoreOnce"]').attributes('disabled')).toBeUndefined()
    const none = await mountSheet(memorySource(book()))
    expect(api(none).checkSpelling()).toBe(false)
  })
})
