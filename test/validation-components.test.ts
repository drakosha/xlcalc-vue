// The components of data validation on their own: the dropdown's arrow and list, the input message and its place, the
// alert's buttons by style with Excel's default and texts, Esc and the answers.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import ValidationAlert from '../src/components/ValidationAlert.vue'
import ValidationDropdown from '../src/components/ValidationDropdown.vue'
import ValidationPrompt from '../src/components/ValidationPrompt.vue'
import { ruMessages } from '../src'
import type { SheetValidationAlert } from '../src/source'
import { settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const CELL = { top: 40, left: 100, width: 64, height: 20 }
const ITEMS = ['Да', 'Нет', '', 'Может быть']
const style = (w: VueWrapper, sel: string) => (w.find(sel).element as HTMLElement).style

describe('ValidationDropdown', () => {
  it('puts the arrow right of the cell and opens nothing until asked', async () => {
    const w = mount(ValidationDropdown, { props: { rect: CELL, items: ITEMS, open: false }, attachTo: document.body })
    expect([style(w, '.xs-vd-anchor').left, style(w, '.xs-vd-anchor').top, style(w, '.xs-vd-anchor').height]).toEqual(['164px', '40px', '20px'])
    expect(w.find('[data-xs-dropdown-arrow]').attributes('aria-expanded')).toBe('false')
    expect(w.find('[data-xs-dropdown-list]').exists()).toBe(false)
    await w.find('[data-xs-dropdown-arrow]').trigger('click')
    expect(w.emitted('toggle')).toHaveLength(1)
  })

  it('lists the items under the cell as wide as the cell and its arrow, a blank one kept; a click gives the item’s text', async () => {
    const w = mount(ValidationDropdown, { props: { rect: CELL, items: ITEMS, open: true, active: 1, messages: ruMessages }, attachTo: document.body })
    const list = w.find('[data-xs-dropdown-list]')
    expect([style(w, '.xs-vd-list').left, style(w, '.xs-vd-list').top, style(w, '.xs-vd-list').width]).toEqual(['100px', '60px', '81px'])
    expect(list.findAll('[role="option"]').map((o) => o.text())).toEqual(['Да', 'Нет', '', 'Может быть'])
    expect(list.find('[aria-selected="true"]').text()).toBe('Нет')
    expect(w.find('[data-xs-dropdown-arrow]').attributes('title')).toBe('Выбрать из списка (Alt+↓)')
    await w.find('[data-xs-dropdown-item="3"]').trigger('click')
    expect(w.emitted('pick')).toEqual([[3, 'Может быть']])
    await w.find('[data-xs-dropdown-item="0"]').trigger('pointermove')
    expect(w.emitted('hover')).toEqual([[0]])
  })

  it('opens the list over the cell where the view has no room below; a finger gets 44 px', async () => {
    const view = { top: 0, left: 0, width: 390, height: 120 }
    const w = mount(ValidationDropdown, { props: { rect: CELL, items: ITEMS, open: true, view, touch: true }, attachTo: document.body })
    // four items of 44 px and the border do not fit under the cell's bottom at 60 px, nor over its top at 40 px: below
    expect(style(w, '.xs-vd-list').top).toBe('60px')
    expect(w.classes()).toContain('xs-vd-touch')
    expect(style(w, '.xs-vd-list').width).toBe('108px')
    await w.setProps({ rect: { ...CELL, top: 300 }, view: { ...view, height: 340 } })
    expect(style(w, '.xs-vd-list').top).toBe(`${300 - 4 * 44 - 2}px`)
  })

  it('closes on a press outside while open, not on one inside', async () => {
    const w = mount(ValidationDropdown, { props: { rect: CELL, items: ITEMS, open: true }, attachTo: document.body })
    w.find('[data-xs-dropdown-item="1"]').element.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    expect(w.emitted('close')).toBeUndefined()
    document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    expect(w.emitted('close')).toHaveLength(1)
    await w.setProps({ open: false })
    document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    expect(w.emitted('close')).toHaveLength(1)
  })
})

describe('ValidationPrompt', () => {
  it('shows the title in bold and the text below the cell, from its middle', async () => {
    const w = mount(ValidationPrompt, { props: { rect: CELL, prompt: { title: 'Ответ', text: 'Выберите из списка' } }, attachTo: document.body })
    await settle()
    expect(w.find('b').text()).toBe('Ответ')
    expect(w.find('.xs-vp-text').text()).toBe('Выберите из списка')
    expect([style(w, '.xs-vp').left, style(w, '.xs-vp').top]).toEqual(['132px', '64px'])
    const bare = mount(ValidationPrompt, { props: { rect: CELL, prompt: { text: 'Только текст' } }, attachTo: document.body })
    expect(bare.find('b').exists()).toBe(false)
  })

  it('goes left of the cell’s middle where the view ends on the right, and over the cell where it ends below', async () => {
    const w = mount(ValidationPrompt, {
      props: { rect: { top: 100, left: 300, width: 64, height: 20 }, prompt: { text: 'Длинная подсказка' }, view: { top: 0, left: 0, width: 390, height: 140 } },
      attachTo: document.body,
    })
    Object.defineProperty(w.element, 'offsetWidth', { value: 150 })
    Object.defineProperty(w.element, 'offsetHeight', { value: 40 })
    await w.setProps({ prompt: { text: 'Длинная подсказка' } })
    await settle()
    // the middle at 332 + 150 passes 390: the box ends at the middle; 124 + 40 passes 140: over the cell
    expect([style(w, '.xs-vp').left, style(w, '.xs-vp').top]).toEqual(['182px', '56px'])
  })
})

describe('ValidationAlert', () => {
  const alert = (severity: SheetValidationAlert['severity'], more: Partial<SheetValidationAlert> = {}): SheetValidationAlert => ({ severity, message: 'От 1 до 10', sheet: 'S', cell: 'B2', ...more })
  const buttons = (w: VueWrapper) => w.findAll('[data-xs-alert-button]').map((b) => b.text())
  const focused = () => (document.activeElement as HTMLElement | null)?.dataset.xsAlertButton

  it('stop: the title, the message, Retry and Cancel, open on Retry', async () => {
    const w = mount(ValidationAlert, { props: { alert: alert('stop', { title: 'Число' }) }, attachTo: document.body })
    await settle()
    expect(w.find('[data-xs-alert="stop"]').attributes('role')).toBe('alertdialog')
    expect(w.find('.xs-va-title').text()).toBe('Число')
    expect(w.find('[data-xs-alert-message]').text()).toBe('От 1 до 10')
    expect(w.find('[data-xs-alert-question]').exists()).toBe(false)
    expect(buttons(w)).toEqual(['Retry', 'Cancel'])
    expect(focused()).toBe('retry')
    await w.find('[data-xs-alert-button="retry"]').trigger('click')
    await w.find('[data-xs-alert-button="cancel"]').trigger('click')
    expect(w.emitted('answer')).toEqual([['retry'], ['cancel']])
  })

  it('names Retry by the sheet’s own `retry` text', async () => {
    const w = mount(ValidationAlert, { props: { alert: alert('stop'), messages: { retry: 'Again' } }, attachTo: document.body })
    await settle()
    expect(buttons(w)).toEqual(['Again', 'Cancel'])
  })

  it('warning in Russian: «Microsoft Excel», the message and «Продолжить?», Да, Нет и Отмена, open on Нет', async () => {
    const w = mount(ValidationAlert, { props: { alert: alert('warning'), messages: ruMessages }, attachTo: document.body })
    await settle()
    expect(w.find('.xs-va-title').text()).toBe('Microsoft Excel')
    expect(w.find('[data-xs-alert-question]').text()).toBe('Продолжить?')
    expect(buttons(w)).toEqual(['Да', 'Нет', 'Отмена'])
    expect(w.find('svg').attributes('aria-label')).toBe('Предупреждение')
    expect(focused()).toBe('no')
    for (const b of ['yes', 'no', 'cancel']) await w.find(`[data-xs-alert-button="${b}"]`).trigger('click')
    expect(w.emitted('answer')).toEqual([['accept'], ['retry'], ['cancel']])
  })

  it('information: OK and Cancel, open on OK; Esc is Cancel; Tab goes round the buttons', async () => {
    const w = mount(ValidationAlert, { props: { alert: alert('information') }, attachTo: document.body })
    await settle()
    expect(buttons(w)).toEqual(['OK', 'Cancel'])
    expect(focused()).toBe('ok')
    await w.find('.xs-va').trigger('keydown', { key: 'Tab' })
    expect(focused()).toBe('cancel')
    await w.find('.xs-va').trigger('keydown', { key: 'Tab' })
    expect(focused()).toBe('ok')
    await w.find('[data-xs-alert-button="ok"]').trigger('click')
    await w.find('.xs-va').trigger('keydown', { key: 'Escape' })
    expect(w.emitted('answer')).toEqual([['accept'], ['cancel']])
    await w.setProps({ alert: null })
    expect(w.find('[data-xs-alert-back]').exists()).toBe(false)
  })
})
