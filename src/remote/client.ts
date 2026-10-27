// A sheet source over a workbook on a server that serves it with `serveSource`: each method is a POST of its arguments,
// and the source's events come over one stream of server-sent events, read with fetch so that the host's headers go too.
import type { EditOutcome, SourceEvent } from '../contract/source'
import { CALLS, isCall, type CallBody, type CallKind, type CallResponse, type Manifest, type StreamEvent } from './methods'
import type { RemoteSource, RemoteSourceOptions } from './types'

/** The longest wait before the event stream connects again. */
const RETRY_MAX = 30_000

class HttpError extends Error {
  constructor(
    readonly status: number,
    message: string,
  ) {
    super(message)
  }
}

const isAbort = (e: unknown) => e instanceof DOMException && e.name === 'AbortError'

/**
 * Connects to a workbook served at `url` (the path before `/manifest`, `/call` and `/events`): reads what the server
 * allows this viewer and gives a source with those methods. Rejects when the server cannot be reached or refuses.
 */
export async function remoteSource(url: string, opts: RemoteSourceOptions = {}): Promise<RemoteSource> {
  const base = url.replace(/\/+$/, '')
  const fetcher = opts.fetch ?? ((input: RequestInfo | URL, init?: RequestInit) => fetch(input, init))
  const headers = () => new Headers(typeof opts.headers === 'function' ? opts.headers() : opts.headers)

  async function request(path: string, init: RequestInit = {}): Promise<Response> {
    const h = headers()
    if (init.body) h.set('Content-Type', 'application/json')
    const res = await fetcher(`${base}/${path}`, { ...init, headers: h })
    if (!res.ok) {
      const body = (await res.json().catch(() => null)) as { error?: string } | null
      throw new HttpError(res.status, body?.error || `${res.status} ${res.statusText}`)
    }
    return res
  }

  const manifest = (await (await request('manifest')).json()) as Manifest
  let history = manifest.history

  /** A method's call: its signal stays here and stops the request; an outcome method fails as an outcome. */
  function method(name: string): (...args: unknown[]) => Promise<unknown> {
    const kind: CallKind = isCall(name) ? CALLS[name] : {}
    return async (...args) => {
      let signal: AbortSignal | undefined
      const sent = [...args]
      if (kind.signal !== undefined) {
        const at = sent[kind.signal] as { signal?: AbortSignal } | AbortSignal | undefined
        if (kind.inOpts && at && !(at instanceof AbortSignal)) {
          const { signal: s, ...rest } = at
          signal = s
          sent[kind.signal] = rest
        } else {
          signal = at as AbortSignal | undefined
          sent[kind.signal] = undefined
        }
      }
      const body: CallBody = { method: name, args: sent, undef: sent.flatMap((a, i) => (a === undefined ? [i] : [])) }
      try {
        const got = (await (await request('call', { method: 'POST', body: JSON.stringify(body), signal })).json()) as Extract<CallResponse, { result: unknown }>
        if (got.history) history = got.history
        return got.result ?? undefined
      } catch (e) {
        if (!kind.outcome || isAbort(e)) throw e
        const out: EditOutcome = { ok: false, reason: e instanceof HttpError && e.status === 403 ? 'refused' : 'failed', message: e instanceof Error ? e.message : String(e) }
        return name === 'replace' ? { ...out, replaced: 0 } : out
      }
    }
  }

  // The event stream: one for every listener, opened by the first and closed with the last or by `close`
  const listeners = new Set<(event: SourceEvent) => void>()
  let stream: AbortController | null = null
  let stopped = false
  const deliver = (event: SourceEvent) => {
    for (const l of [...listeners]) l(event)
  }

  async function connect(ctl: AbortController): Promise<void> {
    let wait = opts.retryMs ?? 1000
    let broke = false
    while (!ctl.signal.aborted) {
      try {
        const res = await request('events', { signal: ctl.signal })
        // Events may have been missed while the stream was broken
        if (broke) deliver({ kind: 'reload' })
        wait = opts.retryMs ?? 1000
        await read(res, ctl.signal)
      } catch (e) {
        // A refusal (the workbook closed, the viewer's access gone) ends the stream for good
        if (ctl.signal.aborted || (e instanceof HttpError && e.status >= 400 && e.status < 500)) return
      }
      broke = true
      if (ctl.signal.aborted) return
      await new Promise((r) => setTimeout(r, wait))
      wait = Math.min(wait * 2, RETRY_MAX)
    }
  }

  /** Reads server-sent events until the stream ends: each `data` is a `StreamEvent`; comments keep the stream open. */
  async function read(res: Response, signal: AbortSignal): Promise<void> {
    const reader = res.body!.pipeThrough(new TextDecoderStream()).getReader()
    signal.addEventListener('abort', () => void reader.cancel().catch(() => {}), { once: true })
    let buf = ''
    for (;;) {
      const { value, done } = await reader.read()
      if (done) return
      buf += value
      let at: number
      while ((at = buf.indexOf('\n\n')) >= 0) {
        const block = buf.slice(0, at)
        buf = buf.slice(at + 2)
        const data = block
          .split('\n')
          .filter((line) => line.startsWith('data:'))
          .map((line) => line.slice(5).trimStart())
          .join('\n')
        if (!data) continue
        const got = JSON.parse(data) as StreamEvent
        if (got.history) history = got.history
        deliver(got.event)
      }
    }
  }

  const source: Record<string, unknown> = {}
  for (const name of manifest.methods) if (isCall(name)) source[name] = method(name)
  if (manifest.methods.includes('subscribe')) {
    source.subscribe = (listener: (event: SourceEvent) => void) => {
      listeners.add(listener)
      if (!stream && !stopped) void connect((stream = new AbortController()))
      return () => {
        listeners.delete(listener)
        if (!listeners.size) {
          stream?.abort()
          stream = null
        }
      }
    }
  }
  return Object.defineProperties(source, {
    access: { value: manifest.access, enumerable: true },
    ...(manifest.locale ? { locale: { value: manifest.locale, enumerable: true } } : {}),
    ...(manifest.maxCells !== undefined ? { maxCells: { value: manifest.maxCells, enumerable: true } } : {}),
    ...(manifest.autoSumFunctions ? { autoSumFunctions: { value: manifest.autoSumFunctions, enumerable: true } } : {}),
    ...(history ? { history: { get: () => history, enumerable: true } } : {}),
    close: {
      value: () => {
        stopped = true
        stream?.abort()
        stream = null
      },
    },
  }) as unknown as RemoteSource
}
