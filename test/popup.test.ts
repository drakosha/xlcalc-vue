// What the boxes over the grid share: a place under the field or over it, kept inside the view; a list's height and its
// scroll to the highlighted item; a press outside while one is open.
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import { clampLeft, listHeight, placeBelowOrAbove, scrollToItem, useOutsidePress } from '../src/components/popup'

afterEach(() => {
  document.body.innerHTML = ''
})

const FIELD = { top: 100, left: 50, width: 80, height: 20 }
const VIEW = { top: 0, left: 0, width: 300, height: 200 }

describe('placeBelowOrAbove', () => {
  it('puts the box under the field, over it only where the view has no room below and room above', () => {
    expect(placeBelowOrAbove(FIELD, 60, VIEW)).toBe(120)
    expect(placeBelowOrAbove(FIELD, 90, VIEW)).toBe(10)
    // room neither way: under it
    expect(placeBelowOrAbove(FIELD, 110, VIEW)).toBe(120)
    // no view: under it
    expect(placeBelowOrAbove(FIELD, 500, undefined)).toBe(120)
  })

  it('keeps a gap from the field both ways', () => {
    expect(placeBelowOrAbove(FIELD, 60, VIEW, 4)).toBe(124)
    expect(placeBelowOrAbove(FIELD, 80, VIEW, 4)).toBe(16)
  })
})

describe('clampLeft and listHeight', () => {
  it('moves a box left to the view’s right edge, not past its left one', () => {
    expect(clampLeft(50, 100, VIEW)).toBe(50)
    expect(clampLeft(250, 100, VIEW)).toBe(200)
    expect(clampLeft(250, 400, VIEW)).toBe(0)
    expect(clampLeft(250, 400, undefined)).toBe(250)
  })

  it('counts up to the rows shown, of 18 px or 44 under a finger, and the border', () => {
    expect([listHeight(3, 8, false), listHeight(20, 8, false), listHeight(3, 8, true)]).toEqual([56, 146, 134])
  })
})

describe('scrollToItem', () => {
  /** A list 100 px high of ten items of 30 px, the layout given by hand. */
  function list(): HTMLElement {
    const ul = document.createElement('ul')
    Object.defineProperty(ul, 'clientHeight', { value: 100 })
    for (let i = 0; i < 10; i++) {
      const li = document.createElement('li')
      Object.defineProperties(li, { offsetTop: { value: i * 30 }, offsetHeight: { value: 30 } })
      ul.append(li)
    }
    return ul
  }

  it('scrolls the list until the item is in view, by as little as it takes', () => {
    const ul = list()
    scrollToItem(ul, 5)
    expect(ul.scrollTop).toBe(80)
    scrollToItem(ul, 3)
    expect(ul.scrollTop).toBe(80)
    scrollToItem(ul, 1)
    expect(ul.scrollTop).toBe(30)
    scrollToItem(ul, -1)
    scrollToItem(null, 2)
    expect(ul.scrollTop).toBe(30)
  })
})

describe('useOutsidePress', () => {
  it('calls back on a press outside the element while active, not inside it, not once inactive', async () => {
    const open = ref(true)
    let calls = 0
    const Box = defineComponent(() => {
      const el = ref<HTMLElement | null>(null)
      useOutsidePress(el, () => calls++, () => open.value)
      return () => h('div', { ref: el, class: 'box' }, [h('span', { class: 'in' })])
    })
    const w = mount(Box, { attachTo: document.body })
    const press = (el: Element) => el.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    press(w.find('.in').element)
    press(document.body)
    expect(calls).toBe(1)
    open.value = false
    await w.vm.$nextTick()
    press(document.body)
    expect(calls).toBe(1)
    open.value = true
    await w.vm.$nextTick()
    w.unmount()
    press(document.body)
    expect(calls).toBe(1)
  })

  it('takes a selector of what counts as inside', () => {
    let calls = 0
    const Box = defineComponent(() => {
      useOutsidePress('[data-mine]', () => calls++, () => true)
      return () => h('div', [h('b', { 'data-mine': '' }, [h('i')]), h('u')])
    })
    const w = mount(Box, { attachTo: document.body })
    w.find('i').element.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    w.find('u').element.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    expect(calls).toBe(1)
  })
})
