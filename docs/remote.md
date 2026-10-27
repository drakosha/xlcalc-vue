# A workbook on a server

A workbook that is calculated on the host's server reaches the sheet in the browser through two halves of the package. On the server, `serveSource` from `@xlcalc/vue/server` serves any `SheetSource` (usually a `localSource`) as a handler of standard requests and responses, which the host mounts under a path of its own server. In the browser, `remoteSource` from `@xlcalc/vue/remote` connects to that path and gives the sheet a `SheetSource` whose methods are calls to the server. Neither half needs the engine or Vue at runtime; neither knows the host's users, its routes or its framework.

One source serves every viewer of a workbook: their calls go to it in turn, and its events, the other viewers' changes among them, go to all of them, so an edit made by one viewer appears on the others' sheets at once.

## The server

`serveSource(source, { access })` returns a `SourceServer`, whose `handle(req)` answers four routes at the end of any path: `GET …/manifest` (what the viewer may call and the source's readonly fields), `POST …/call` (a method's call), `POST …/batch` (several calls that read, together) and `GET …/events` (the source's events as server-sent events). The host keeps one server per workbook and decides who reaches it. `access` is `read` (the default) or `write`, the same for every viewer, or a function of the request that returns it for the viewer the request comes from. A viewer with `read` gets only the methods that do not change the workbook, and its sheet is read-only; a call that would change the workbook is refused with 403 whatever the client sends.

With [Hono](https://hono.dev), as an example:

```ts
import { Hono } from 'hono'
import { createCalcPool } from 'xlcalc/worker'
import { localSource } from '@xlcalc/vue/local'
import { serveSource, type SourceServer } from '@xlcalc/vue/server'

const calc = createCalcPool({ threads: 2 })
const books = new Map<string, SourceServer>()

async function bookServer(id: string): Promise<SourceServer> {
  let server = books.get(id)
  if (!server) {
    const source = localSource(await loadBook(id), { calc })
    server = serveSource(source, { access: (req) => (canEdit(req, id) ? 'write' : 'read') })
    books.set(id, server)
  }
  return server
}

const app = new Hono()
app.all('/sheet/:id/*', async (c) => (await bookServer(c.req.param('id'))).handle(c.req.raw))
```

`loadBook` and `canEdit` are the host's: where its workbooks live and who may change them. On Node, `localSource` calculates on the server's own thread unless `calc` says otherwise; a server with several workbooks passes them one pool of xlcalc's threads. Any server that hands a `Request` to a function and sends back its `Response` works the same way (Bun, Deno, Node with an adapter such as `@hono/node-server`).

When the workbook is replaced (a newer version, the file written by someone else), `close(message)` sends every viewer `reload` with the message and ends their event streams; later requests to that server get 410. The source itself stays open: the host closes it (`localSource`'s `close()`) and serves the new one, under the same path or another.

## The browser

```vue
<script setup lang="ts">
import { onUnmounted, shallowRef } from 'vue'
import { SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '@xlcalc/vue'
import { remoteSource, type RemoteSource } from '@xlcalc/vue/remote'
import '@xlcalc/vue/style.css'

const props = defineProps<{ id: string }>()
const source = shallowRef<RemoteSource | null>(null)
const failed = shallowRef('')
remoteSource(`/sheet/${props.id}`).then(
  (s) => (source.value = s),
  (e: Error) => (failed.value = e.message),
)
onUnmounted(() => source.value?.close())

const state = shallowRef<SheetViewState | null>(null)
const sheet = shallowRef<SheetViewApi | null>(null)
</script>

<template>
  <p v-if="failed">{{ failed }}</p>
  <template v-else-if="source">
    <SheetToolbar :source="source" :sheet="sheet" :state="state" :locale="source.locale" />
    <SheetView ref="sheet" :source="source" :locale="source.locale" style="height: 420px" @state="state = $event" />
  </template>
</template>
```

`remoteSource(url, opts)` reads the manifest first and resolves with a source that has exactly the methods the server allows this viewer, with `access`, the workbook's `locale` (when the served source has one, as `localSource` does), `maxCells`, `autoSumFunctions` and `history`. It rejects when the server cannot be reached or refuses the viewer. `fetch` replaces the global fetch, and `headers` (a value, or a function called for each request, for a token that changes) go with every request, the event stream's too: the stream is read with fetch rather than `EventSource`, which cannot send headers.

A method that returns an `EditOutcome` and fails on the server (an error, a network failure, 403, 410) resolves with one that is not `ok`, as the sheet displays any rejection: `refused` for 403, `failed` otherwise, with the server's message. Any other method, `learnWord` included, rejects with an `Error` carrying the message. Calls that read and are made in the same turn go as one request (`batch`), whose response gives each call's answer as the server completes it, so a slow window holds back no other call: Insert Function's hundreds of signatures, or a screen of windows, are one request. A method that changes the workbook always goes alone, in its order. A call's `AbortSignal` stops it, as the sheet stops the windows it no longer needs (in a batch it fails alone, and the request stops when all its calls have); the server passes the request's own signal to the methods that take one, so a `localSource` stops its reads too.

The sheet and the toolbar subscribe to one event stream between them, opened by the first listener and closed with the last, or by `close()`. The server counts the source's events, and a stream starts by telling the count: when it differs from the count the client last saw (in the manifest or the last event), the workbook changed while the client was not listening, before its first stream opened or while a stream was broken, and the listeners get `reload`. When the stream breaks it connects again, after a second and then twice as long each time up to 30 s; a refusal (4xx, such as 410 after the server's `close`) ends it for good. Each tab keeps one stream open, which counts against a browser's six connections to a host over HTTP/1.1: serve the sheet over HTTP/2 when a viewer may open several workbooks at once.

## Several viewers

- **One history.** The served source has one undo history for the workbook, as `localSource` has: Undo takes back the workbook's last change, whoever made it, and every viewer's toolbar shows the same steps (`history` comes with each call's response and each event). Excel keeps each co-author's own changes apart; this does not, yet.
- **Edits in turn.** Calls that change the workbook run one after another in the order they reach the source; two viewers typing into the same cell leave the later value, and each sees the other's change as it arrives.
- **Each viewer's own view.** What only the viewer sees stays in the browser: the selection, the display filter and its sort, the zoom, the widths kept by `storageKey`. A `localSource` keeps the rows of each display filter apart (up to eight a sheet), so viewers filtering one sheet differently each keep their rows until they apply their filter again ([local.md](local.md#display-filter-and-totals)).
- **What every viewer shares.** Everything in the workbook, including the Spelling dictionary's added words (`learnWord`), which is a change of the served source.

## The protocol

For a host that writes its own client or tests the server directly: the manifest is `{ protocol: 1, access, methods, maxCells?, autoSumFunctions?, history?, locale?, epoch?, seq? }`, where `methods` lists the callable methods and `subscribe` when there are events, and `epoch` and `seq` are the server's mark and the number of the source's events so far. A call's body is `{ method, args, undef? }`: the method's arguments in order, without its signal, and in `undef` the indexes of those that were `undefined`, which JSON would turn into `null`. A batch's body is `{ calls }`, 1 to 1,000 call bodies of methods that do not change the workbook; its response is newline-delimited JSON, a line `{ i, status, result?, history?, error? }` for each call as it completes, `i` its index in `calls`. A call's response is `{ result?, history? }` (no `result` for a method that returns nothing), or `{ error }` with status 400 (not a call of a method of the contract), 403 (read-only viewer), 404 (a method the source does not have), 410 (closed) or 500 (the method threw). Each message of the stream is `data:` with `{ epoch, seq, event?, history? }`: the first has no event and gives the count the stream starts from, the others carry the source's `SourceEvent` and the count after it; lines starting with `:` keep the stream open (every 25 s, `keepAliveMs`).
