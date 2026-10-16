// Excel's tip with a function's arguments while a formula is typed (`SUMIFS(sum_range, criteria_range1, …)`): the call the
// caret stands in, its signature asked of the source once per name, the argument the caret is at in bold.
import { computed, getCurrentScope, onScopeDispose, shallowRef, watch, type Ref } from 'vue'
import type { ScanOptions } from 'xlcalc/contract'
import { callAt } from 'xlcalc/text'
import type { SheetFunctionSignature, SheetSource } from '../source'
import type { EditorState } from './editing'

/** A piece of the tip's text; the argument the caret is at is `current`. */
export interface TipPart {
  text: string
  current?: boolean
}

/**
 * The innermost function call around `caret` in formula text typed in `opts.locale` (`callAt` of `xlcalc/text`): the name
 * the source is asked by, English in upper case (`SUM` for a typed `сумм`), and the argument's index from 0; null outside one.
 */
export function callUnder(text: string, caret: number, opts: ScanOptions): { name: string; index: number } | null {
  const call = callAt(text, caret, opts)
  return call && { name: call.fn || call.name, index: call.index }
}

/** An argument's name past the list: the one `repeat` before it with its number counted on (`number2` → `number3`). */
function nameAfter(names: string[], i: number, repeat: number): string {
  if (i < names.length) return names[i]!
  return nameAfter(names, i - repeat, repeat).replace(/\d+$/, (n) => String(Number(n) + 1))
}

/**
 * The tip in pieces with the argument at `index` current; past the list, the repeated group goes on with its numbers
 * counted on until it holds the argument (`SUM(number1, [number2], [number3], ...)`), as Excel's tip. A text that does
 * not hold the names in order is shown whole, nothing current.
 */
export function tipParts(sig: SheetFunctionSignature, index: number): TipPart[] {
  const { text, args, repeat } = sig
  const spans: { at: number; end: number }[] = []
  let pos = text.indexOf('(') + 1
  for (const a of args) {
    const at = text.indexOf(a.name, pos)
    if (!pos || at < 0) return [{ text }]
    spans.push({ at, end: at + a.name.length })
    pos = at + a.name.length
  }
  // The repeated groups past the list go in before the closing `sep...)`, each in brackets as the last one is.
  const tail = text.lastIndexOf('...)')
  const close = text.lastIndexOf(']', tail)
  const extend = index >= args.length && repeat > 0 && tail > 0 && close > 0 && args.length >= repeat
  const parts: TipPart[] = []
  let at = 0
  spans.forEach((s, i) => {
    parts.push({ text: text.slice(at, s.at) }, { text: text.slice(s.at, s.end), ...(i === index ? { current: true } : {}) })
    at = s.end
  })
  if (extend) {
    const sep = text.slice(close + 1, tail)
    parts.push({ text: text.slice(at, close + 1) })
    const names = args.map((a) => a.name)
    for (let first = args.length; first <= index; first += repeat) {
      parts.push({ text: `${sep}[` })
      for (let i = first; i < first + repeat; i++) {
        if (i > first) parts.push({ text: sep })
        parts.push({ text: nameAfter(names, i, repeat), ...(i === index ? { current: true } : {}) })
      }
      parts.push({ text: ']' })
    }
    at = close + 1
  }
  parts.push({ text: text.slice(at) })
  return parts.filter((p) => p.text)
}

export interface FunctionTipDeps {
  source: () => SheetSource
  editor: Ref<EditorState | null>
  /** The open field's element, for its caret. */
  input: () => HTMLTextAreaElement | null
  /** How the book's formula text reads; null while its locale loads. */
  scan: () => ScanOptions | null
  /** The list of functions is open: it goes first, the tip waits under it. */
  listOpen: () => boolean
}

/** The tip's state over the editor's draft: the signature of the call at the caret and its argument, asked once per name. */
export function useFunctionTip(deps: FunctionTipDeps) {
  /** The signature shown and the argument the caret is at. */
  const tip = shallowRef<{ sig: SheetFunctionSignature; index: number } | null>(null)
  const parts = computed(() => (tip.value && !deps.listOpen() ? tipParts(tip.value.sig, tip.value.index) : null))
  /** Signatures by the name's upper case for one source; null for a name it does not know. */
  let cache = new Map<string, SheetFunctionSignature | null>()
  let cacheOf: SheetSource | null = null
  /** The call the caret stands in now, by the name's key. */
  let want: { key: string; index: number } | null = null
  let pending: { key: string; ctrl: AbortController } | null = null
  let timer: ReturnType<typeof setTimeout> | null = null
  /** What the tip was last made for: a key's follow after the input's update finds the same and does nothing. */
  let asked: { text: string; caret: number | null; src: SheetSource; opts: ScanOptions | null } | null = null

  function drop(): void {
    pending?.ctrl.abort()
    pending = null
  }
  function close(): void {
    drop()
    want = null
    asked = null
    tip.value = null
  }
  /** The tip for the caret at `caret` in `text` (null when text is selected): from the cache, else asked of the source by name. */
  function update(text: string, caret: number | null): void {
    const src = deps.source()
    const opts = deps.scan()
    if (asked && asked.text === text && asked.caret === caret && asked.src === src && asked.opts === opts) return
    const at = caret !== null && opts && src.functionSignature ? callUnder(text, caret, opts) : null
    if (!at) return close()
    asked = { text, caret, src, opts }
    if (cacheOf !== src) {
      cache = new Map()
      cacheOf = src
    }
    const key = at.name.toUpperCase()
    want = { key, index: at.index }
    if (cache.has(key)) {
      drop()
      const sig = cache.get(key)
      tip.value = sig ? { sig, index: at.index } : null
      return
    }
    if (pending?.key === key) return
    drop()
    tip.value = null
    const entry = { key, ctrl: new AbortController() }
    pending = entry
    Promise.resolve()
      .then(() => src.functionSignature!(at.name, entry.ctrl.signal))
      .then(
        (sig) => {
          if (pending !== entry) return
          pending = null
          cache.set(key, sig)
          if (want?.key === key && sig) tip.value = { sig, index: want.index }
        },
        () => {
          // A failed request is not kept: the name is asked again when the caret comes to it next.
          if (pending === entry) {
            pending = null
            asked = null
          }
        },
      )
  }
  /** The caret moved in the field (a key, a click): the tip follows it once the move is done. */
  function follow(): void {
    if (timer) clearTimeout(timer)
    timer = setTimeout(() => {
      timer = null
      const el = deps.input()
      const ed = deps.editor.value
      if (el && ed) update(ed.text, el.selectionStart === el.selectionEnd ? el.selectionStart : null)
      else close()
    })
  }

  /** A function's signature for Insert Function, from the tip's cache or asked of the source and kept; null for a name the source does not know or without `functionSignature`. */
  async function signature(name: string): Promise<SheetFunctionSignature | null> {
    const src = deps.source()
    if (!src.functionSignature) return null
    if (cacheOf !== src) {
      cache = new Map()
      cacheOf = src
    }
    const key = name.toUpperCase()
    if (cache.has(key)) return cache.get(key)!
    // A failed request is not kept, as the tip's.
    const sig = await src.functionSignature(name).catch(() => undefined)
    if (sig === undefined) return null
    if (cacheOf === src) cache.set(key, sig)
    return sig
  }

  // A field opened, moved between the cell and the bar or closed (Esc, a commit), or a new source: the tip goes.
  watch([() => deps.editor.value, () => deps.editor.value?.place, deps.source], close)
  if (getCurrentScope())
    onScopeDispose(() => {
      drop()
      if (timer) clearTimeout(timer)
    })

  return { parts, update, follow, close, signature }
}

export type FunctionTip = ReturnType<typeof useFunctionTip>
