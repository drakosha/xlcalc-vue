// The `@xlcalc/vue/server` entry point: a sheet source served to viewers over HTTP, as a handler of standard requests and
// responses that any server mounts (Hono, Node with an adapter, Bun, Deno). One source serves every viewer of a workbook:
// their calls go to it in turn, and its events, the viewers' changes included, go to all of them.
import type { SheetSource, SourceEvent } from '../contract/source'
import { BATCH_MAX, CALLS, historyOf, isCall, type BatchLine, type CallBody, type CallResponse, type Manifest, type StreamEvent } from './methods'
import type { ServeOptions, SourceAccess, SourceServer } from './types'

export type { ServeOptions, SourceAccess, SourceServer } from './types'

const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' } })
const fail = (status: number, error: string) => json({ error } satisfies CallResponse, status)
const messageOf = (e: unknown) => (e instanceof Error ? e.message : String(e))

/** Serves `source` to viewers: the host keeps one per workbook and closes the source itself when it is done. */
export function serveSource(source: SheetSource, opts: ServeOptions = {}): SourceServer {
  const accessOf = async (req: Request): Promise<SourceAccess> => (typeof opts.access === 'function' ? await opts.access(req) : (opts.access ?? 'read'))
  /** The open event streams, ended by `close`. */
  const streams = new Set<{ send: (e: SourceEvent) => void; end: () => void }>()
  let closed: string | null = null
  const gone = () => fail(410, closed || 'The workbook was closed')
  /** The source's events so far, under this server's own mark: a client that sees another count missed some. */
  const epoch = Math.random().toString(36).slice(2)
  let seq = 0
  const counting = source.subscribe?.call(source, () => seq++)

  function manifest(access: SourceAccess): Manifest {
    const methods: string[] = Object.keys(CALLS).filter((m) => isCall(m) && typeof source[m] === 'function' && (access === 'write' || !CALLS[m].write))
    if (source.subscribe) methods.push('subscribe')
    const locale = (source as { locale?: unknown }).locale
    const dateSystem = (source as { dateSystem?: unknown }).dateSystem
    return {
      protocol: 1,
      access,
      methods,
      ...(source.maxCells !== undefined ? { maxCells: source.maxCells } : {}),
      ...(source.autoSumFunctions ? { autoSumFunctions: source.autoSumFunctions } : {}),
      ...(source.history ? { history: historyOf(source) } : {}),
      ...(typeof locale === 'string' ? { locale } : {}),
      ...(dateSystem === 1900 || dateSystem === 1904 ? { dateSystem } : {}),
      ...(source.subscribe ? { epoch, seq } : {}),
    }
  }

  /** A call's status and response; `reads` refuses a method that changes the workbook, as a batch must. */
  async function run(body: Partial<CallBody> | null, req: Request, access: SourceAccess, reads = false): Promise<{ status: number; out: CallResponse }> {
    const refuse = (status: number, error: string) => ({ status, out: { error } })
    const method = body?.method
    if (typeof method !== 'string' || !isCall(method) || !Array.isArray(body?.args)) return refuse(400, `No method ${String(method)}`)
    const fn = source[method] as ((...args: unknown[]) => Promise<unknown>) | undefined
    if (!fn) return refuse(404, `The source has no ${method}`)
    const kind = CALLS[method]
    if (kind.write && access !== 'write') return refuse(403, 'The workbook is read-only for this viewer')
    if (kind.write && reads) return refuse(400, `A batch takes no ${method}: a change goes alone, in its order`)
    const args = [...body.args]
    for (const i of Array.isArray(body.undef) ? body.undef : []) if (Number.isInteger(i)) args[i] = undefined
    if (kind.signal !== undefined) args[kind.signal] = kind.inOpts ? { ...(args[kind.signal] as object | undefined), signal: req.signal } : req.signal
    try {
      const result = await fn.apply(source, args)
      return { status: 200, out: { ...(result === undefined ? {} : { result }), ...(source.history ? { history: historyOf(source) } : {}) } }
    } catch (e) {
      return refuse(500, messageOf(e))
    }
  }

  /** The request's JSON, or null when it is not JSON. */
  const bodyOf = (req: Request): Promise<unknown> => req.json().catch(() => undefined)

  async function call(req: Request, access: SourceAccess): Promise<Response> {
    const body = await bodyOf(req)
    if (body === undefined) return fail(400, 'The call is not JSON')
    if (closed !== null) return gone()
    const { status, out } = await run(body as Partial<CallBody> | null, req, access)
    return json(out, status)
  }

  /** Calls that read, run together: each line of the response is one call's index, status and response, as it completes. */
  async function batch(req: Request, access: SourceAccess): Promise<Response> {
    const body = (await bodyOf(req)) as { calls?: unknown } | null | undefined
    if (closed !== null) return gone()
    const calls = body?.calls
    if (!Array.isArray(calls) || !calls.length || calls.length > BATCH_MAX) return fail(400, `A batch is 1 to ${BATCH_MAX} calls`)
    const encoder = new TextEncoder()
    const stream = new ReadableStream<Uint8Array>({
      async start(controller) {
        await Promise.all(
          calls.map(async (c: Partial<CallBody> | null, i) => {
            const { status, out } = await run(c, req, access, true)
            controller.enqueue(encoder.encode(`${JSON.stringify({ i, status, ...out } satisfies BatchLine)}\n`))
          }),
        )
        controller.close()
      },
    })
    return new Response(stream, { headers: { 'Content-Type': 'application/x-ndjson', 'Cache-Control': 'no-store', 'X-Accel-Buffering': 'no' } })
  }

  function events(req: Request): Response {
    const subscribe = source.subscribe
    if (!subscribe) return fail(404, 'The source has no events')
    const encoder = new TextEncoder()
    let stop = () => {}
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        const write = (text: string) => controller.enqueue(encoder.encode(text))
        const data = (body: StreamEvent) => write(`data: ${JSON.stringify(body)}\n\n`)
        const entry = {
          send(event: SourceEvent) {
            data({ event, epoch, seq, ...(source.history ? { history: historyOf(source) } : {}) })
          },
          // The viewer gone or the server closed: the stream ends, so that whoever reads it stops too
          end() {
            stop()
            controller.close()
          },
        }
        const unsubscribe = subscribe.call(source, entry.send)
        const alive = setInterval(() => write(': alive\n\n'), opts.keepAliveMs ?? 25_000)
        streams.add(entry)
        let stopped = false
        stop = () => {
          if (stopped) return
          stopped = true
          clearInterval(alive)
          unsubscribe()
          streams.delete(entry)
          req.signal.removeEventListener('abort', entry.end)
        }
        req.signal.addEventListener('abort', entry.end)
        // The count the stream starts from, after its subscription: the client compares it with the count it knows
        data({ epoch, seq })
      },
      cancel: () => stop(),
    })
    return new Response(stream, { headers: { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-store', 'X-Accel-Buffering': 'no' } })
  }

  return {
    async handle(req) {
      if (closed !== null) return gone()
      const end = new URL(req.url).pathname.split('/').pop()
      const access = await accessOf(req)
      // The host may have closed the server while the viewer's access was being decided
      if (closed !== null) return gone()
      if (end === 'manifest' && req.method === 'GET') return json(manifest(access))
      if (end === 'call' && req.method === 'POST') return call(req, access)
      if (end === 'batch' && req.method === 'POST') return batch(req, access)
      if (end === 'events' && req.method === 'GET') return events(req)
      return fail(404, 'Not found')
    },
    close(message) {
      closed = message ?? ''
      counting?.()
      for (const s of [...streams]) {
        s.send({ kind: 'reload', ...(message ? { message } : {}) })
        s.end()
      }
    },
  }
}
