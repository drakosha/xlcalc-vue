// Test helpers: mounting the sheet, settling promises, and a source whose answers the test controls.
import { flushPromises, mount, type VueWrapper } from '@vue/test-utils'
import { nextTick } from 'vue'
import type { CellValue } from 'xlcalc/contract'
import { SheetView } from '../src'
import type { CellEntry, EditOutcome, SheetInfo, SheetSource, SheetWindow, SourceEvent, WindowRequest } from '../src/source'
import { cellAddress } from '../src/core/a1'

/** Let promises, watchers and renders run out. */
export async function settle(rounds = 6): Promise<void> {
  for (let i = 0; i < rounds; i++) {
    await flushPromises()
    await nextTick()
  }
}

export async function mountSheet(source: SheetSource, props: Record<string, unknown> = {}): Promise<VueWrapper> {
  const wrapper = mount(SheetView, { props: { source, ...props }, attachTo: document.body })
  await settle()
  return wrapper
}

/** A window of a rectangle with text from `cell` (values: numbers when the text is one). */
export function gridWindow(req: WindowRequest, cell: (row: number, col: number) => string, extra: Partial<SheetWindow> = {}): SheetWindow {
  const text: string[][] = []
  const values: CellValue[][] = []
  for (let r = req.row; r < req.row + req.rows; r++) {
    const t: string[] = []
    const v: CellValue[] = []
    for (let c = req.col; c < req.col + req.cols; c++) {
      const s = cell(r, c)
      t.push(s)
      v.push(s === '' ? null : /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : s)
    }
    text.push(t)
    values.push(v)
  }
  return { row: req.row, col: req.col, text, values, formulas: {}, colors: {}, styleIds: {}, styles: [], widths: new Array(req.cols).fill(null), rowHeights: {}, merges: [], ...extra }
}

interface Pending {
  req: WindowRequest
  resolve: (w: SheetWindow) => void
  reject: (e: unknown) => void
}

/** A source with sheets of the given size; `manual` holds window answers until the test releases them. */
export function fakeSource(sheets: SheetInfo[], opts: { manual?: boolean; cell?: (sheet: string, row: number, col: number) => string; edit?: (entries: CellEntry[]) => Promise<EditOutcome> } = {}) {
  const cell = opts.cell ?? ((sheet: string, r: number, c: number) => `${sheet}:${cellAddress(r, c)}`)
  const requests: WindowRequest[] = []
  const held: Pending[] = []
  const edits: CellEntry[][] = []
  const listeners = new Set<(e: SourceEvent) => void>()
  let sheetsCalls = 0
  const build = (req: WindowRequest) => {
    const info = sheets.find((s) => s.name === req.sheet)!
    const rows = Math.max(0, Math.min(req.rows, info.maxRow - req.row + 1))
    const cols = Math.max(0, Math.min(req.cols, info.maxCol - req.col + 1))
    return gridWindow({ ...req, rows, cols }, (r, c) => cell(req.sheet, r, c))
  }
  const source: SheetSource = {
    sheets: async () => {
      sheetsCalls++
      return sheets
    },
    window: (req) => {
      requests.push(req)
      if (!opts.manual) return Promise.resolve(build(req))
      return new Promise((resolve, reject) => held.push({ req, resolve, reject }))
    },
    edit: async (entries) => {
      edits.push(entries)
      return opts.edit ? opts.edit(entries) : { ok: true, changes: [] }
    },
    subscribe: (l) => {
      listeners.add(l)
      return () => listeners.delete(l)
    },
  }
  return {
    source,
    requests,
    held,
    edits,
    get sheetsCalls() {
      return sheetsCalls
    },
    /** Answer a held request with its rectangle. */
    release(p: Pending) {
      held.splice(held.indexOf(p), 1)
      p.resolve(build(p.req))
    },
    emit(e: SourceEvent) {
      for (const l of listeners) l(e)
    },
  }
}

/** The text of a drawn cell, by address. */
export function cellText(wrapper: VueWrapper, row: number, col: number): string | undefined {
  const td = wrapper.find(`[data-cell="${row}:${col}"]`)
  return td.exists() ? td.text() : undefined
}
