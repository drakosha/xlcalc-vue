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
  /** The server's mark and the count of the source's events this client has seen, from the manifest on. */
  let seen = { epoch: manifest.epoch, seq: manifest.seq }

  /** A method's call: its signal stays here and stops the request; an outcome method fails as an outcome. */
  function method(name: string): (...args: unknown[]) => Promise<unknown> {
    const kind: CallKind = isCall(name) ? CALLS[name] : {}
    return async (...args) => {
      let signal: AbortSignal | undefined
      // As long as the signal's place, without holes: a hole would arrive as null
      const sent = Array.from({ length: Math.max(args.length, (kind.signal ?? -1) + 1) }, (_, i) => args[i])
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
        const got = (await (await request('call', { method: 'POST', body: JSON.stringify(body), signal })).json()) as Exclude<CallResponse, { error: string }>
        if (got.history) history = got.history
        return got.result
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
    while (!ctl.signal.aborted) {
      try {
        const res = await request('events', { signal: ctl.signal })
        wait = opts.retryMs ?? 1000
        await read(res, ctl.signal)
      } catch (e) {
        // A refusal (the workbook closed, the viewer's access gone) ends the stream for good
        if (ctl.signal.aborted || (e instanceof HttpError && e.status >= 400 && e.status < 500)) return
      }
      if (ctl.signal.aborted) return
      await new Promise((r) => setTimeout(r, wait))
      wait = Math.min(wait * 2, RETRY_MAX)
    }
  }

  /** A message of the stream: its count first, which tells whether events were missed before it opened (`reload`), then events. */
  function take(got: StreamEvent): void {
    if (got.history) history = got.history
    const missed = !got.event && (got.epoch !== seen.epoch || got.seq !== seen.seq)
    seen = { epoch: got.epoch, seq: got.seq }
    if (got.event) deliver(got.event)
    else if (missed) deliver({ kind: 'reload' })
  }

  /** Reads server-sent events until the stream ends: each `data` is a `StreamEvent`; comments keep the stream open. */
  async function read(res: Response, signal: AbortSignal): Promise<void> {
    const reader = res.body!.pipeThrough(new TextDecoderStream()).getReader()
    const cancel = () => void reader.cancel().catch(() => {})
    signal.addEventListener('abort', cancel, { once: true })
    try {
      await readEvents(reader)
    } finally {
      signal.removeEventListener('abort', cancel)
    }
  }

  async function readEvents(reader: ReadableStreamDefaultReader<string>): Promise<void> {
    let buf = ''
    // A CR at a chunk's end may be the first half of a CRLF
    let cr = ''
    for (;;) {
      const { value, done } = await reader.read()
      if (done) return
      let text = cr + value
      cr = text.endsWith('\r') ? '\r' : ''
      if (cr) text = text.slice(0, -1)
      buf += text.replace(/\r\n?/g, '\n')
      let at: number
      while ((at = buf.indexOf('\n\n')) >= 0) {
        const block = buf.slice(0, at)
        buf = buf.slice(at + 2)
        const data = block
          .split('\n')
          .filter((line) => line.startsWith('data:'))
          .map((line) => line.slice(5).trimStart())
          .join('\n')
        if (data) take(JSON.parse(data) as StreamEvent)
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
