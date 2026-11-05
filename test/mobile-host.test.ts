// Hosts can collapse their chrome during editing without reaching into the sheet's DOM.
import { afterEach, expect, it, vi } from 'vitest'
import { memorySource } from '../src'
import type { SheetViewState } from '../src/contract/view'
import { fingerPointer, mountSheet, settle } from './helpers'

afterEach(() => { document.body.innerHTML = ''; vi.restoreAllMocks() })

it('reports editing until a draft is committed or cancelled', async () => {
  const w = await mountSheet(memorySource({ sheets: [{ name: 'One', cells: { A1: 'original' } }] }), { compact: true })
  const state = () => w.emitted('state')!.at(-1)![0] as SheetViewState
  expect(w.classes()).toContain('xs-compact')
  expect(state().editing).toBeUndefined()
  await w.find('.xs-fx-input').trigger('focus')
  await settle()
  expect(state().editing).toBe(true)
  await w.find('.xs-fx-input').setValue('draft')
  await w.find('[data-xs-bar-cancel]').trigger('click')
  await settle()
  expect(state().editing).toBeUndefined()
  expect(w.find('[data-cell="1:1"]').text()).toBe('original')
  await w.find('.xs-fx-input').trigger('focus')
  await w.find('.xs-fx-input').setValue('saved')
  await w.find('[data-xs-bar-enter]').trigger('click')
  await settle()
  expect(state().editing).toBeUndefined()
  expect(w.find('[data-cell="1:1"]').text()).toBe('saved')
  w.unmount()
})

it('opens a touch menu without preselecting Cut and preserves keyboard navigation', async () => {
  fingerPointer()
  const w = await mountSheet(memorySource({ sheets: [{ name: 'One', cells: { A1: 'original' } }] }), { compact: true })
  await w.find('[data-cell="1:1"]').trigger('contextmenu', { clientX: 60, clientY: 90 })
  await settle()
  const menu = w.find('[data-xs-menu-box]')
  expect(menu.classes()).toContain('xs-menu-mobile')
  expect(document.activeElement).toBe(menu.element)
  expect(w.find('[data-xs-menu="cut"]').element).not.toBe(document.activeElement)
  await menu.trigger('keydown', { key: 'ArrowDown' })
  expect(document.activeElement?.getAttribute('role')).toBe('menuitem')
  await menu.trigger('keydown', { key: 'Escape' })
  await settle()
  expect(w.find('[data-xs-menu-box]').exists()).toBe(false)
  w.unmount()
})
