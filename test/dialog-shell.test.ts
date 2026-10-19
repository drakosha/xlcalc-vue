// The sheet's modal box: focus moved inside on opening, Esc and a click on the backdrop close it (the latter only where
// allowed), and Tab past the last control wraps to the first and Shift+Tab back.
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { h } from 'vue'
import DialogShell from '../src/components/DialogShell.vue'
import { settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const content = () => [h('input', { class: 'first', value: 'abc' }), h('button', { class: 'middle' }), h('button', { class: 'last' }), h('button', { class: 'off', disabled: true })]
const focused = () => (document.activeElement as HTMLElement | null)?.className

describe('DialogShell', () => {
  it('focuses the control its selector names on opening, an input with its text selected; the box itself without one', async () => {
    const w = mount(DialogShell, { props: { focus: '.first' }, slots: { default: content }, attachTo: document.body })
    await settle()
    const input = w.find('.first').element as HTMLInputElement
    expect(document.activeElement).toBe(input)
    expect([input.selectionStart, input.selectionEnd]).toEqual([0, 3])
    const bare = mount(DialogShell, { slots: { default: content }, attachTo: document.body })
    await settle()
    expect(document.activeElement).toBe(bare.find('.xs-modal').element)
  })

  it('passes on what it is given to the box: its class, label and listeners, with its own role and aria-modal', async () => {
    const w = mount(DialogShell, { props: { role: 'alertdialog' }, attrs: { class: 'mine', 'aria-label': 'Box', 'data-x': '1' }, slots: { default: content }, attachTo: document.body })
    const box = w.find('.xs-modal')
    expect([box.classes('mine'), box.attributes('aria-label'), box.attributes('data-x'), box.attributes('role'), box.attributes('aria-modal')]).toEqual([true, 'Box', '1', 'alertdialog', 'true'])
    expect(w.find('.xs-modal-back').classes('mine')).toBe(false)
  })

  it('closes on Esc, and on a click on the backdrop unless not allowed', async () => {
    const w = mount(DialogShell, { slots: { default: content }, attachTo: document.body })
    await w.find('.middle').trigger('keydown', { key: 'Escape' })
    await w.find('.xs-modal-back').trigger('pointerdown')
    await w.find('.middle').trigger('pointerdown')
    expect(w.emitted('close')).toHaveLength(2)
    const alert = mount(DialogShell, { props: { outside: false }, slots: { default: content }, attachTo: document.body })
    await alert.find('.xs-modal-back').trigger('pointerdown')
    expect(alert.emitted('close')).toBeUndefined()
  })

  it('keeps Tab inside: past the last enabled control to the first, Shift+Tab from the first to the last', async () => {
    const w = mount(DialogShell, { props: { focus: '.last' }, slots: { default: content }, attachTo: document.body })
    await settle()
    expect(focused()).toBe('last')
    await w.find('.last').trigger('keydown', { key: 'Tab' })
    expect(focused()).toBe('first')
    await w.find('.first').trigger('keydown', { key: 'Tab', shiftKey: true })
    expect(focused()).toBe('last')
    // between the two the browser's own order applies
    ;(w.find('.middle').element as HTMLElement).focus()
    const e = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
    w.find('.middle').element.dispatchEvent(e)
    expect([e.defaultPrevented, focused()]).toEqual([false, 'middle'])
  })
})
