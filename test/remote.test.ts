// @vitest-environment node
// The transport over a real HTTP server: one local source served to two viewers, read and write access, an edit by one
// viewer reaching the other, the history, errors as outcomes, a request's signal, and the server closing.
import type { Server } from 'node:http'
import { afterEach, describe, expect, it } from 'vitest'
import type { SheetFunctionSignature, SheetSource, SourceEvent } from '../src/contract/source'
import { localSource, type LocalSource } from '../src/local'
import { remoteSource, type RemoteSource } from '../src/remote'
import { serveSource, type ServeOptions, type SourceServer } from '../src/remote/server'
import { listen } from '../scripts/node-handler'

const open: { server?: Server; local?: LocalSource; clients: RemoteSource[] } = { clients: [] }
afterEach(async () => {
  for (const c of open.clients) c.close()
  open.clients = []
  open.server?.closeAllConnections()
  open.server?.close()
  await open.local?.close()
  open.server = open.local = undefined
})

/** A ru-RU workbook served with `opts`: the server, the source and the address of its routes. */
async function served(opts: ServeOptions = { access: 'write' }, source?: SheetSource): Promise<{ serving: SourceServer; url: string }> {
  const s = source ?? (open.local = localSource({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 2 }, A2: { value: 3 }, A3: { formula: 'SUM(A1:A2)' } } }] }, { calc: 'inline' }))
  const serving = serveSource(s, opts)
  const { server, url } = await listen((req) => serving.handle(req))
  open.server = server
  return { serving, url: `${url}/book/1` }
}
const connect = async (url: string, opts = {}) => {
  const c = await remoteSource(url, opts)
  open.clients.push(c)
  return c
}
/** The events a source's subscriber receives, and a promise of the next one that matches. */
function events(s: SheetSource) {
  const got: SourceEvent[] = []
  const waiting: { test: (e: SourceEvent) => boolean; done: (e: SourceEvent) => void }[] = []
  s.subscribe!((e) => {
    got.push(e)
    for (const w of [...waiting]) if (w.test(e)) (waiting.splice(waiting.indexOf(w), 1), w.done(e))
  })
  return { got, next: (test: (e: SourceEvent) => boolean = () => true) => new Promise<SourceEvent>((done) => waiting.push({ test, done })) }
}

describe('remoteSource over serveSource', () => {
  it('reads the sheets and windows of the served workbook, with its locale, maxCells and history', async () => {
    const { url } = await served()
    const c = await connect(url)
    expect([c.access, c.locale, c.maxCells, c.history]).toEqual(['write', 'ru-RU', open.local!.maxCells, { undo: 0, redo: 0 }])
    expect((await c.sheets()).map((x) => x.name)).toEqual(['Лист1'])
    const w = await c.window({ sheet: 'Лист1', row: 1, rows: 3, col: 1, cols: 1 })
    expect('text' in w && w.text.map((r) => r[0])).toEqual(['2', '3', '5'])
    expect('formulas' in w && w.formulas.A3).toBe('СУММ(A1:A2)')
    // An argument left out stays left out, rather than becoming null
    expect((await c.stats!('Лист1', 'A1:A3', undefined)).sum).toBe(10)
    expect((await c.autoSum!('Лист1', 'A4'))?.formula).toBe('СУММ(A3)')
  })

  it("gives a read-only viewer only the methods that do not change the workbook, and refuses a write it sends anyway", async () => {
    const { url } = await served({ access: (req) => (req.headers.get('x-user') === 'editor' ? 'write' : 'read') })
    const reader = await connect(url)
    const editor = await connect(url, { headers: { 'x-user': 'editor' } })
    expect(reader.access).toBe('read')
    expect([reader.edit, reader.undo, reader.format, reader.learnWord]).toEqual([undefined, undefined, undefined, undefined])
    expect([typeof reader.find, typeof reader.subscribe, typeof editor.edit, typeof editor.undo]).toEqual(['function', 'function', 'function', 'function'])
    const res = await fetch(`${url}/call`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ method: 'edit', args: [[{ sheet: 'Лист1', cell: 'A1', text: '9' }]] }) })
    expect([res.status, await res.json()]).toEqual([403, { error: 'The workbook is read-only for this viewer' }])
  })

  it("sends one viewer's edit to every other viewer with the shared history, and its undo too", async () => {
    const { url } = await served()
    const a = await connect(url)
    const b = await connect(url)
    const seen = events(b)
    await new Promise((r) => setTimeout(r, 50))
    const changed = seen.next((e) => e.kind === 'changed')
    const out = await a.edit!([{ sheet: 'Лист1', cell: 'A1', text: '10' }])
    expect(out.ok && out.changes.find((x) => x.cell === 'A3')?.text).toBe('13')
    const e = await changed
    expect(e.kind === 'changed' && e.changes.map((x) => [x.cell, x.text])).toEqual(expect.arrayContaining([['A1', '10'], ['A3', '13']]))
    // One history for the workbook: the other viewer's undo takes back the first one's edit
    expect([a.history, b.history]).toEqual([{ undo: 1, redo: 0 }, { undo: 1, redo: 0 }])
    const undone = seen.next((x) => x.kind === 'changed')
    expect((await b.undo!()).ok).toBe(true)
    expect((await undone).kind === 'changed' && b.history).toEqual({ undo: 0, redo: 1 })
    const w = await a.window({ sheet: 'Лист1', row: 1, rows: 3, col: 1, cols: 1 })
    expect('text' in w && w.text.map((r) => r[0])).toEqual(['2', '3', '5'])
  })

  it('reloads a viewer whose stream opened after a change it may have missed, and only then', async () => {
    const { url } = await served()
    const late = await connect(url)
    const quiet = await connect(url)
    const other = await connect(url)
    const quietEvents = events(quiet)
    await new Promise((r) => setTimeout(r, 50))
    // The change comes after `late` read the workbook and before its stream opens
    expect((await other.edit!([{ sheet: 'Лист1', cell: 'A1', text: '7' }])).ok).toBe(true)
    await quietEvents.next((e) => e.kind === 'changed')
    expect(await events(late).next()).toEqual({ kind: 'reload' })
    expect(quietEvents.got.map((e) => e.kind)).not.toContain('reload')
  })

  it("returns a failure of a method that changes the workbook as an outcome, and rejects with the server's message otherwise", async () => {
    const broken: SheetSource = {
      sheets: async () => {
        throw new Error('No disk')
      },
      window: async () => ({ pending: { elapsedMs: 5 } }),
      edit: async () => {
        throw new Error('Locked by the agent')
      },
    }
    const { url } = await served({ access: 'write' }, broken)
    const c = await connect(url)
    await expect(c.sheets()).rejects.toThrow('No disk')
    expect(await c.window({ sheet: 'S', row: 1, rows: 1, col: 1, cols: 1 })).toEqual({ pending: { elapsedMs: 5 } })
    expect(await c.edit!([{ sheet: 'S', cell: 'A1', text: '1' }])).toEqual({ ok: false, reason: 'failed', message: 'Locked by the agent' })
    expect(c.subscribe).toBeUndefined()
  })

  it("passes an argument left out as left out and null as null, and a signal as the request's", async () => {
    const seen: unknown[][] = []
    const stub: SheetSource = {
      sheets: async () => [],
      window: async () => ({ pending: { elapsedMs: 0 } }),
      stats: async (...args) => (seen.push(args), { cells: 0, numbers: 0, sum: 0, min: null, max: null }),
      link: async (...args) => (seen.push(args), { ok: true, changes: [] }),
      validation: async () => null,
    }
    const { url } = await served({ access: 'write' }, stub)
    const c = await connect(url)
    await c.stats!('S', 'A1:A3', undefined)
    await c.link!('S', 'A1', null)
    await c.stats!('S', 'B1:B2')
    expect(await c.validation!('S', 'A1')).toBeNull()
    expect(seen[0]!.slice(0, 3)).toEqual(['S', 'A1:A3', undefined])
    expect(seen[0]![3]).toBeInstanceOf(AbortSignal)
    expect(seen[1]).toEqual(['S', 'A1', null])
    // A call shorter than its signal's place leaves no hole that would arrive as null
    expect(seen[2]!.slice(0, 3)).toEqual(['S', 'B1:B2', undefined])
  })

  it("sends a turn's reads as one request, answered call by call as each completes, and a change alone", async () => {
    let requests = 0
    let release!: () => void
    const held = new Promise<void>((r) => (release = r))
    const stub: SheetSource = {
      sheets: async () => [],
      window: async () => ({ pending: { elapsedMs: 0 } }),
      functionSignature: async (name) => {
        if (name === 'SLOW') await held
        if (name === 'BAD') throw new Error('No such function here')
        return { name, args: [] } as unknown as SheetFunctionSignature
      },
      edit: async () => ({ ok: true, changes: [] }),
    }
    const { url } = await served({ access: () => (requests++, 'write') }, stub)
    const c = await connect(url)
    const names = Array.from({ length: 300 }, (_, i) => `F${i}`)
    const got = await Promise.all(names.map((n) => c.functionSignature!(n)))
    expect(got.map((x) => x?.name)).toEqual(names)
    expect(requests).toBe(2)
    // In one batch: the slow call holds back no other, a failed or stopped call fails alone
    const stop = new AbortController()
    const slow = c.functionSignature!('SLOW')
    const failure = (p: Promise<unknown>) => p.then(() => null, (e: Error) => e)
    const bad = failure(c.functionSignature!('BAD'))
    const stopped = failure(c.functionSignature!('F1', stop.signal))
    const fast = c.functionSignature!('F2')
    const edited = c.edit!([{ sheet: 'S', cell: 'A1', text: '1' }])
    stop.abort()
    expect((await fast)?.name).toBe('F2')
    expect((await bad)?.message).toBe('No such function here')
    expect((await stopped)?.name).toBe('AbortError')
    expect((await edited).ok).toBe(true)
    release()
    expect((await slow)?.name).toBe('SLOW')
    // the manifest, the 300 reads, these four reads and the edit
    expect(requests).toBe(4)
  })

  it("stops a call with the caller's signal", async () => {
    let started!: () => void
    const begun = new Promise<void>((r) => (started = r))
    const slow: SheetSource = {
      sheets: async () => [],
      window: (_req, signal) =>
        new Promise((_ok, fail) => {
          started()
          signal?.addEventListener('abort', () => fail(signal.reason))
        }),
    }
    const { url } = await served({}, slow)
    const c = await connect(url)
    const stop = new AbortController()
    const call = c.window({ sheet: 'S', row: 1, rows: 1, col: 1, cols: 1 }, stop.signal)
    await begun
    stop.abort()
    await expect(call).rejects.toMatchObject({ name: 'AbortError' })
  })

  it('on close sends every viewer reload with the message and refuses later calls', async () => {
    const { serving, url } = await served()
    const c = await connect(url)
    const seen = events(c)
    await new Promise((r) => setTimeout(r, 50))
    const reload = seen.next()
    serving.close('A newer version of the workbook is open')
    expect(await reload).toEqual({ kind: 'reload', message: 'A newer version of the workbook is open' })
    expect(await c.edit!([{ sheet: 'Лист1', cell: 'A1', text: '1' }])).toEqual({ ok: false, reason: 'failed', message: 'A newer version of the workbook is open' })
  })
})
