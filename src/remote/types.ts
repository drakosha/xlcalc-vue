// Public types of the transport: the client's source over a server (`./remote`) and the server's handler (`./server`).
import type { SheetSource } from '../contract/source'

/** What a viewer may do: `read` calls only the methods that do not change the workbook, `write` every method. */
export type SourceAccess = 'read' | 'write'

export interface RemoteSourceOptions {
  /** The fetch to call the server with, for a host that adds what its server needs; the global fetch without it. */
  fetch?: typeof fetch
  /** Headers of every request, or a function that gives them for each (a token that changes). */
  headers?: HeadersInit | (() => HeadersInit)
  /** The first wait before the event stream connects again after it broke, doubled up to 30 s; 1,000 ms by default. */
  retryMs?: number
}

/** A sheet source over a workbook on a server: the methods the server allows, with its events. */
export interface RemoteSource extends SheetSource {
  /** What the server allows this viewer: with `read` the source has no method that changes the workbook. */
  readonly access: SourceAccess
  /** The workbook's locale, when the server's source has one (as `localSource` does). */
  readonly locale?: string
  /** The workbook's date system, when the server's source has one (as `localSource` does for a workbook in 1904). */
  readonly dateSystem?: 1900 | 1904
  /** Stops the event stream; the source's calls keep working. */
  close(): void
}

export interface ServeOptions {
  /** What each request may do, `read` by default: a value for every viewer, or a function of the request (its user). */
  access?: SourceAccess | ((req: Request) => SourceAccess | Promise<SourceAccess>)
  /** How often the event stream sends a comment, so that proxies keep it open; 25,000 ms by default. */
  keepAliveMs?: number
}

/** A source served to viewers: a handler of standard requests and responses, which any server mounts under a path. */
export interface SourceServer {
  /**
   * Answers a request whose path ends with `/manifest` (GET: what the viewer may call), `/call` (POST: a call), `/batch`
   * (POST: calls that read, answered line by line) or `/events` (GET: the source's events as server-sent events); anything else is 404.
   */
  handle(req: Request): Promise<Response>
  /** Sends every viewer `reload` with the message and ends their streams; later requests get 410. The source stays open. */
  close(message?: string): void
}
