<script setup lang="ts">
// The sheet over localSource: xlcalc computes the synthetic book in a Web Worker (`?calc=inline` in the page's thread);
// timings of the opening and of each edit are shown and kept in `window.__xs` for the browser check.
import { onBeforeUnmount, ref, shallowRef } from 'vue'
import { SheetView, ruMessages, type SheetViewState } from '../src'
import { localSource, saveXlsx } from '../src/local'
import type { CellEntry } from '../src/source'
import { localBook } from './local-book'

const params = new URLSearchParams(location.search)
const calc = params.get('calc') === 'inline' ? 'inline' : 'pool'
const t0 = performance.now()
const book = localBook(Number(params.get('rows') ?? 3400))
const built = performance.now() - t0
const local = localSource(book, { calc })
const timings = ref({ built, ready: 0, edit: 0, changes: 0 })
const failed = ref('')
const focus = shallowRef(params.get('sheet') ? { ref: params.get('sheet')! } : null)
const state = ref<SheetViewState | null>(null)

const debug = { timings: timings.value, edits: [] as { ms: number; changes: number }[], save: () => saveXlsx(local).then((x) => x.data.byteLength) }
;(window as unknown as { __xs: typeof debug }).__xs = debug
local.ready.then(
  () => (timings.value.ready = performance.now() - t0),
  (e: unknown) => (failed.value = e instanceof Error ? e.message : String(e)),
)

/** The source with its edits timed: from the call to the outcome, recomputation included. */
const source = {
  ...local,
  async edit(entries: CellEntry[]) {
    const start = performance.now()
    const got = await local.edit(entries)
    const ms = performance.now() - start
    const changes = got.ok ? got.changes.length : 0
    timings.value.edit = ms
    timings.value.changes = changes
    debug.edits.push({ ms, changes })
    return got
  },
}

async function save(): Promise<void> {
  const { data } = await saveXlsx(local)
  const a = document.createElement('a')
  a.href = URL.createObjectURL(new Blob([data as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
  a.download = 'synthetic.xlsx'
  a.click()
  URL.revokeObjectURL(a.href)
}
onBeforeUnmount(() => void local.close())
const ms = (n: number) => `${Math.round(n)} мс`
</script>

<template>
  <div class="page">
    <header class="top">
      <strong>xlcalc-sheet/local</strong>
      <span class="state" data-calc>{{ calc === 'pool' ? 'Web Worker' : 'тот же поток' }}</span>
      <span class="state" data-timings>
        книга {{ ms(timings.built) }} · расчёт {{ timings.ready ? ms(timings.ready) : '…' }}<template v-if="timings.edit"> · правка {{ ms(timings.edit) }}, ячеек {{ timings.changes }}</template>
      </span>
      <span v-if="failed" class="error">{{ failed }}</span>
      <button type="button" @click="save">Сохранить xlsx</button>
      <span class="state">{{ state ? `${state.sheet}!${state.ref}${state.formula ? ' =' + state.formula : ''}` : '' }}</span>
    </header>
    <SheetView class="sheet" :source="source" :messages="ruMessages" :locale="local.locale" :focus="focus" storage-key="xs-local" @state="state = $event" />
  </div>
</template>

<style>
.page {
  display: flex;
  flex-direction: column;
  height: 100%;
  font: 13px system-ui, sans-serif;
}
.top {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 10px;
  padding: 6px 10px;
  border-bottom: 1px solid #d0d4da;
}
.state {
  color: #6b7280;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.error {
  color: #b91c1c;
}
.sheet {
  flex: 1 1 auto;
  min-height: 0;
}
@media (max-width: 600px) {
  .top .state:last-of-type,
  .top button {
    display: none;
  }
}
</style>
