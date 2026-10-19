// Excel's "We found a typo in your formula" dialog: an unclosed formula that the source rejects with `correction` is
// offered corrected; Yes enters the correction as the same key would, No says what is missing and reopens the editor.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { ruMessages } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { active, address, cellText, editorKey, editorOpen, goTo, key, mountSheet, settle, type, until } from './helpers'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

/** The sheet over an empty ru-RU workbook calculated on the page, with the Russian texts. */
async function mountRu(): Promise<VueWrapper> {
  const source = localSource({ locale: 'ru-RU', sheets: [{ name: 'S', cells: {} }] }, { calc: 'inline' })
  sources.push(source)
  await source.ready
  const w = await mountSheet(source, { messages: ruMessages })
  await settle(12)
  return w
}
const alertOf = (severity: string) => document.querySelector(`[data-xs-alert="${severity}"]`)
const button = (b: string) => document.querySelector<HTMLElement>(`[data-xs-alert-button="${b}"]`)!
const field = (w: VueWrapper) => w.find('.xs-editor').element as HTMLTextAreaElement
/** The sheet's call has been answered: an alert is open, or no cell is waiting for the source any more. */
const answered = () => until(() => !!document.querySelector('[data-xs-alert]') || !document.querySelector('.xs-pending'))
/** A call after Yes is written: the prompt closed and no cell waiting. */
const written = () => until(() => !alertOf('typo') && !document.querySelector('.xs-pending'))
const escape = (severity: string) => alertOf(severity)!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
/** Types `text` over the active cell and finishes it with the editor's key. */
async function enter(w: VueWrapper, text: string, k = 'Enter', extra: Record<string, unknown> = {}): Promise<void> {
  await key(w, text[0]!)
  await type(w, text)
  await editorKey(w, k, extra)
  await settle(12)
  await answered()
}

describe('a formula left open with a correction', () => {
  it('Enter on =(1+2 asks to accept =(1+2), Yes the default; Yes enters it and the selection goes down', async () => {
    const w = await mountRu()
    await enter(w, '=(1+2')
    const box = alertOf('typo')
    expect(box?.querySelector('[data-xs-alert-message]')?.textContent).toBe(ruMessages.typoFound('=(1+2)'))
    expect(box?.querySelector('[data-xs-alert-message]')?.textContent).toBe('Мы обнаружили опечатку в формуле и попытались ее исправить: =(1+2). Принять исправление?')
    expect([...box!.querySelectorAll('[data-xs-alert-button]')].map((b) => b.textContent?.trim())).toEqual(['Да', 'Нет'])
    expect(box?.querySelector('header')?.textContent).toBe('Microsoft Excel')
    // Enter presses the focused default button.
    expect(document.activeElement).toBe(button('yes'))
    expect(editorOpen(w)).toBe(false)
    expect(active(w)).toBe('1:1')
    button('yes').click()
    await settle(12)
    await written()
    expect(alertOf('typo')).toBeNull()
    expect(cellText(w, 1, 1)).toBe('3')
    expect(active(w)).toBe('2:1')
    await goTo(w, 'A1')
    expect(field(w).value).toBe('')
    expect((w.find('.xs-fx-input').element as HTMLTextAreaElement).value).toBe('=(1+2)')
  })

  it('No, or Esc, says a bracket is missing; after its OK the field is open with the text typed', async () => {
    for (const how of ['no', 'escape'] as const) {
      const w = await mountRu()
      await enter(w, '=1+СУММ(2')
      expect(alertOf('typo')?.textContent).toContain('=1+СУММ(2)')
      if (how === 'no') button('no').click()
      else escape('typo')
      await settle()
      expect(alertOf('typo')).toBeNull()
      expect(alertOf('parenthesis')?.querySelector('[data-xs-alert-message]')?.textContent).toBe('В формуле не хватает открывающей или закрывающей круглой скобки.')
      expect([...alertOf('parenthesis')!.querySelectorAll('[data-xs-alert-button]')].map((b) => b.textContent?.trim())).toEqual(['ОК'])
      expect(document.activeElement).toBe(button('ok'))
      expect(editorOpen(w)).toBe(false)
      button('ok').click()
      await settle()
      expect(alertOf('parenthesis')).toBeNull()
      expect(editorOpen(w)).toBe(true)
      expect(field(w).value).toBe('=1+СУММ(2')
      expect(active(w)).toBe('1:1')
      expect(cellText(w, 1, 1)).not.toBe('3')
      w.unmount()
    }
  })

  it('Tab, Control+Return over a range and the formula bar’s Enter (✓) go the same way after Yes', async () => {
    const w = await mountRu()
    await enter(w, '=(1+2', 'Tab')
    button('yes').click()
    await settle(12)
    await written()
    expect(cellText(w, 1, 1)).toBe('3')
    expect(active(w)).toBe('1:2')

    await goTo(w, 'B2:B3')
    await enter(w, '=(2+2', 'Enter', { ctrlKey: true })
    expect(alertOf('typo')?.textContent).toContain('=(2+2)')
    expect(address(w)).toBe('B2:B3')
    button('yes').click()
    await settle(12)
    await written()
    expect([cellText(w, 2, 2), cellText(w, 3, 2)]).toEqual(['4', '4'])
    expect(address(w)).toBe('B2:B3')

    await goTo(w, 'C1')
    await key(w, '=')
    await type(w, '=(3+3')
    await w.find('[data-xs-bar-enter]').trigger('click')
    await settle(12)
    await answered()
    button('yes').click()
    await settle(12)
    await written()
    expect(cellText(w, 1, 3)).toBe('6')
    expect(active(w)).toBe('1:3')
  })

  it('without a correction the question is Fix or Restore, as before', async () => {
    const w = await mountRu()
    await enter(w, '=СУММ(A1:A2))')
    expect(alertOf('typo')).toBeNull()
    expect(alertOf('formula')).not.toBeNull()
    button('fix').click()
    await settle()
    expect(editorOpen(w)).toBe(true)
    expect(field(w).value).toBe('=СУММ(A1:A2))')
  })
})
