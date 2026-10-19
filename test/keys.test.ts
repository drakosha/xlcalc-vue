// Shortcuts by the key's position on any keyboard layout: Ctrl+A, Ctrl+F and Ctrl+H as Excel handles them, while AltGr (Ctrl+Alt) still types its character.
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { address, editorOpen, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
})

const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'apple', B2: 5 } }] })

describe('shortcut letters on another layout', () => {
  it('Control+A selects all by the key under A: Russian and Greek', async () => {
    for (const k of ['ф', 'α']) {
      const w = await mountSheet(memorySource(book()))
      await key(w, k, { ctrlKey: true, code: 'KeyA' })
      await settle()
      expect(address(w)).toBe('1:1048576')
      w.unmount()
    }
  })

  it('Control+F and Control+H open the search bar by the keys under F and H on a Greek layout', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, 'φ', { ctrlKey: true, code: 'KeyF' })
    await settle()
    expect(w.find('.xs-find').exists()).toBe(true)
    expect(w.find('[data-xs-replace-input]').exists()).toBe(false)
    await key(w, 'Escape')
    await settle()
    await key(w, 'η', { ctrlKey: true, code: 'KeyH' })
    await settle()
    expect(w.find('[data-xs-replace-input]').exists()).toBe(true)
  })

  it('AltGr comes as Control+Alt and types its character: no select all, no search', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, 'ą', { ctrlKey: true, altKey: true, code: 'KeyA' })
    await settle()
    expect(address(w)).toBe('A1')
    expect(editorOpen(w)).toBe(true)
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('ą')
    w.unmount()
    const v = await mountSheet(memorySource(book()))
    await key(v, '[', { ctrlKey: true, altKey: true, code: 'KeyF' })
    await settle()
    expect(v.find('.xs-find').exists()).toBe(false)
    expect((v.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('[')
  })
})
