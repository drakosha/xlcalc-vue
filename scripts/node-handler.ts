// A handler of standard requests and responses on Node's HTTP server, as a host mounts `serveSource` without a framework:
// for the transport's tests and the browser check.
import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http'
import { Readable } from 'node:stream'

/** Node's request handler over `handle`: the request's signal aborts when the connection closes, a stream is sent as it comes. */
export function nodeHandler(handle: (req: Request) => Promise<Response>): (req: IncomingMessage, res: ServerResponse) => void {
  return (req, res) => {
    const abort = new AbortController()
    res.on('close', () => abort.abort())
    void handle(toRequest(req, abort.signal)).then(
      async (out) => {
        res.writeHead(out.status, Object.fromEntries(out.headers))
        res.flushHeaders()
        if (!out.body) return void res.end()
        for await (const chunk of out.body) res.write(chunk)
        res.end()
      },
      (e: unknown) => res.writeHead(500).end(String(e)),
    )
  }
}

/** Serves `handle` on a free port of localhost; resolves with the server and its address. */
export async function listen(handle: (req: Request) => Promise<Response>): Promise<{ server: Server; url: string }> {
  const server = createServer(nodeHandler(handle))
  return { server, url: await started(server) }
}

/** Starts a server on a free port of localhost; resolves with its address. */
export async function started(server: Server): Promise<string> {
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', r))
  const at = server.address()
  return `http://127.0.0.1:${typeof at === 'object' && at ? at.port : 0}`
}

function toRequest(req: IncomingMessage, signal: AbortSignal): Request {
  const body = req.method === 'GET' || req.method === 'HEAD' ? undefined : (Readable.toWeb(req) as ReadableStream<Uint8Array>)
  const headers = new Headers()
  for (const [k, v] of Object.entries(req.headers)) if (typeof v === 'string') headers.set(k, v)
  return new Request(`http://${req.headers.host}${req.url}`, { method: req.method, headers, body, signal, ...(body ? { duplex: 'half' } : {}) } as RequestInit)
}
