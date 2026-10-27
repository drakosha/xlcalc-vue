// Test helpers: mounting the sheet, settling promises, and a source whose responses the test controls.
import { flushPromises, mount, type VueWrapper } from '@vue/test-utils'
import { nextTick } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { CellValue } from 'xlcalc/contract'
import { SheetView } from '../src'
import type { CellEntry, EditOutcome, SheetInfo, SheetSource, SheetWindow, SourceEvent, WindowRequest } from '../src/contract/source'

/** Lets promises, watchers and renders run to completion. */
export async function settle(rounds = 6): Promise<void> {
  for (let i = 0; i < rounds; i++) {
    await flushPromises()
    await nextTick()
  }
}

/**
 * Settles until `ready` holds: for work that takes as many rounds as the machine's load requires (a workbook calculated in
 * slices, a module loaded), where a fixed number of rounds fails on a busy machine. Real time, so it works under fake timers too.
 */
export async function until(ready: () => boolean, ms = 15_000): Promise<void> {
  const end = performance.now() + ms
  while (!ready()) {
    if (performance.now() > end) throw new Error(`until: not ready after ${ms} ms`)
    await settle(1)
  }
}

export async function mountSheet(source: SheetSource, props: Record<string, unknown> = {}): Promise<VueWrapper> {
  const wrapper = mount(SheetView, { props: { source, ...props }, attachTo: document.body })
  await settle()
  return wrapper
}

/** A window over a range with text from `cell` (values: numbers when the text is one). */
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
  signal?: AbortSignal
  resolve: (w: SheetWindow) => void
  reject: (e: unknown) => void
}

/** A source with sheets of the given size; `manual` holds window responses until the test releases them. */
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
    window: (req, signal) => {
      requests.push(req)
      if (!opts.manual) return Promise.resolve(build(req))
      return new Promise((resolve, reject) => held.push({ req, signal, resolve, reject }))
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
    /** Answers a held request with its range. */
    release(p: Pending) {
      held.splice(held.indexOf(p), 1)
      p.resolve(build(p.req))
    },
    emit(e: SourceEvent) {
      for (const l of listeners) l(e)
    },
  }
}

/** Copy or cut on the grid as the browser sends it; returns the text the sheet put on the clipboard. */
export function clipboard(w: VueWrapper, kind: 'copy' | 'cut'): string | undefined {
  const data: Record<string, string> = {}
  const e = new Event(kind, { bubbles: true, cancelable: true }) as ClipboardEvent
  Object.defineProperty(e, 'clipboardData', { value: { setData: (t: string, v: string) => (data[t] = v) } })
  w.find('.xs-scroller').element.dispatchEvent(e)
  return data['text/plain']
}

/** The text of a drawn cell, by address. */
export function cellText(wrapper: VueWrapper, row: number, col: number): string | undefined {
  const td = wrapper.find(`[data-cell="${row}:${col}"]`)
  return td.exists() ? td.text() : undefined
}

// Keys on the grid and in the cell editor, the active cell, the editor's state and the Name Box.
export const key = (w: VueWrapper, k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
export const editorKey = (w: VueWrapper, k: string, extra: Record<string, unknown> = {}) => w.find('.xs-editor').trigger('keydown', { key: k, ...extra })
export const active = (w: VueWrapper) => w.find('.xs-cell.xs-active').attributes('data-cell')
export const editorOpen = (w: VueWrapper) => !w.find('.xs-editor').classes().includes('xs-editor-idle')
export const address = (w: VueWrapper) => (w.find('.xs-address').element as HTMLInputElement).value
export async function type(w: VueWrapper, text: string): Promise<void> {
  const el = w.find('.xs-editor')
  ;(el.element as HTMLTextAreaElement).value = text
  await el.trigger('input')
}
export async function goTo(w: VueWrapper, place: string): Promise<void> {
  const input = w.find('.xs-address')
  await input.setValue(place)
  await input.trigger('keydown', { key: 'Enter' })
  await settle()
}
/** Pastes text on the grid as the browser sends it. */
export function paste(w: VueWrapper, text: string): void {
  const e = new Event('paste', { bubbles: true, cancelable: true }) as ClipboardEvent
  Object.defineProperty(e, 'clipboardData', { value: { getData: () => text } })
  w.find('.xs-scroller').element.dispatchEvent(e)
}
/** The formula bar's input. */
export const barField = (w: VueWrapper) => w.find('.xs-fx-input').element as HTMLTextAreaElement
/** A press on the formula bar's input as the browser sends it: the caret (or a dragged selection) where the mouse put it, then mousedown, focus and click, with `keys` held. */
export async function pressBar(w: VueWrapper, from: number, to = from, keys: MouseEventInit = {}): Promise<void> {
  const el = barField(w)
  el.setSelectionRange(from, to)
  el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, ...keys }))
  el.focus()
  el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, ...keys }))
  el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, ...keys }))
  await settle()
}
