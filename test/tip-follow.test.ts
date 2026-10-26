// A key in the editor: the input event updates the function list and the tooltip, and the key's follow-up after it finds the same draft and caret and does nothing.
import type { VueWrapper } from '@vue/test-utils'
import { expect, it, vi } from 'vitest'
import { watch } from 'vue'
import { functionSignature } from 'xlcalc'
import { memorySource } from '../src'
import { sheetContextKey } from '../src/components/context'
import type { SheetContext } from '../src/core/sheet'
import { editorKey, key, mountSheet, settle, type } from './helpers'

const scans = vi.hoisted(() => ({ n: 0 }))
vi.mock('xlcalc/text', async (load) => {
  const text = await load<typeof import('xlcalc/text')>()
  return { ...text, callAt: (...a: Parameters<typeof text.callAt>) => (scans.n++, text.callAt(...a)) }
})

const contextOf = (w: VueWrapper) => (w.vm.$ as unknown as { provides: Record<symbol, SheetContext> }).provides[sheetContextKey as symbol]!

it('a typed character finds the call at the caret once and sets the list once', async () => {
  const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }], functions: ['SUM', 'SUMIF', 'SUMIFS'] })
  source.functionSignature = async (name) => functionSignature(name)
  const w = await mountSheet(source)
  const ctx = contextOf(w)
  await key(w, '=')
  await type(w, '=SUM(1,S')
  await new Promise((r) => setTimeout(r))
  await settle()
  expect(ctx.ed.ac.open.value).toBe(true)
  let lists = 0
  watch(ctx.ed.ac.items, () => lists++, { flush: 'sync' })
  scans.n = 0
  // A key as the browser sends it: keydown, then input with the new text, then the follow-up's timer.
  await editorKey(w, 'U')
  await type(w, '=SUM(1,SU')
  await new Promise((r) => setTimeout(r))
  await settle()
  expect(scans.n).toBe(1)
  expect(lists).toBe(1)
  expect(ctx.ed.ac.items.value.map((i) => i.text)).toEqual(['SUM', 'SUMIF', 'SUMIFS'])
})
