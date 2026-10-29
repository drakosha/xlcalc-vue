// A split sheet for the tests: a source that keeps the split and has freeze and structure, the lines in the file's points, the panes.
import type { VueWrapper } from '@vue/test-utils'
import { vi } from 'vitest'
import type { SheetViewApi } from '../src'
import type { EditOutcome, SheetInfo, SheetSplit } from '../src/contract/source'
import { colPx } from '../src/core/geometry'
import { fakeSource, settle } from './helpers'

export const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
/** The default row (15 pt) and column on screen, and the headings: the numbers' 56 px and the letters' 24 px. */
export const ROW = 20
export const COL = colPx(null)
export const LEAD = 56
export const HEAD = 24

export function book(split: SheetSplit | undefined, extra: Partial<SheetInfo> = {}, opts: Parameters<typeof fakeSource>[1] = {}) {
  const sheets: SheetInfo[] = [{ name: 'S', maxRow: 300, maxCol: 40, ...(split ? { split } : {}), ...extra }]
  const f = fakeSource(sheets, opts)
  f.source.split = vi.fn(async (_sheet: string, panes: SheetSplit | null): Promise<EditOutcome> => {
    const { split: _, ...rest } = sheets[0]!
    sheets[0] = panes ? { ...rest, split: panes } : rest
    return { ok: true, changes: [] }
  })
  f.source.freeze = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [] }))
  f.source.structure = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [] }))
  return f
}
/** A split after `rows` rows and `cols` columns of the window from A1, in the file's points. */
export const lines = (rows: number, cols: number) => ({ ...(cols ? { x: ((LEAD + cols * COL) * 3) / 4 } : {}), ...(rows ? { y: ((HEAD + rows * ROW) * 3) / 4 } : {}) })
export const pane = (w: VueWrapper, id: number) => w.find(`[data-xs-pane="${id}"]`)
export const el = (w: VueWrapper, id: number) => pane(w, id).element as HTMLElement
export async function scroll(w: VueWrapper, id: number, to: { top?: number; left?: number }): Promise<void> {
  if (to.top !== undefined) el(w, id).scrollTop = to.top
  if (to.left !== undefined) el(w, id).scrollLeft = to.left
  await pane(w, id).trigger('scroll')
  await settle()
}
