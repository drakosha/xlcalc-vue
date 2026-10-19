<script setup lang="ts">
// The sheet over localSource: xlcalc calculates the synthetic workbook, or an xlsx file opened from the file input, in a Web
// Worker (`?calc=inline` on the page's thread); timings of the opening and of each edit are shown and stored in `window.__xs`
// for the browser check, and the workbook saves back to xlsx under its file's name.
import { computed, onBeforeUnmount, ref, shallowRef } from 'vue'
import { SheetView, ruMessages, type SheetViewState } from '../src'
import { groupIssues, localSource, openXlsx, saveXlsx, type IssueGroup, type LocalSource } from '../src/local'
import type { CellEntry, EditOptions } from '../src/contract/source'
import { localBook } from './local-book'

const params = new URLSearchParams(location.search)
const calc = params.get('calc') === 'inline' ? 'inline' : 'pool'
const t0 = performance.now()
const book = localBook(Number(params.get('rows') ?? 3400))
const timings = ref({ built: performance.now() - t0, ready: 0, edit: 0, changes: 0 })
const failed = ref('')
const focus = shallowRef(params.get('sheet') ? { ref: params.get('sheet')! } : null)
const state = ref<SheetViewState | null>(null)
const local = shallowRef<LocalSource>(start(localSource(book, { calc }), t0))
const fileName = ref('synthetic.xlsx')
/** What the opened file contained that the workbook cannot keep, by kind. */
const lost = ref<IssueGroup[]>([])

const debug = { timings: timings.value, edits: [] as { ms: number; changes: number }[], save: () => saveXlsx(local.value).then((x) => x.data.byteLength) }
;(window as unknown as { __xs: typeof debug }).__xs = debug

/** A source whose opening is timed from `since` and whose failure is displayed. */
function start(s: LocalSource, since: number): LocalSource {
  failed.value = ''
  timings.value.ready = 0
  s.ready.then(
    () => (timings.value.ready = performance.now() - since),
    (e: unknown) => (failed.value = e instanceof Error ? e.message : String(e)),
  )
  return s
}

/** The source with its edits timed: from the call to the outcome, recalculation included. */
const source = computed(() => {
  const s = local.value
  return {
    ...s,
    async edit(entries: CellEntry[], opts?: EditOptions) {
      const begun = performance.now()
      const got = await s.edit(entries, opts)
      const ms = performance.now() - begun
      const changes = got.ok ? got.changes.length : 0
      timings.value.edit = ms
      timings.value.changes = changes
      debug.edits.push({ ms, changes })
      return got
    },
  }
})

/** Opens the chosen xlsx file: a new source over its workbook, the old one closed. */
async function openFile(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (!file) return
  const since = performance.now()
  try {
    const read = await openXlsx(file, { locale: 'ru-RU' })
    const previous = local.value
    local.value = start(localSource(read.book, { calc, session: { cached: read.cached } }), since)
    fileName.value = file.name
    lost.value = groupIssues(read.issues)
    focus.value = null
    void previous.close()
  } catch (e) {
    failed.value = e instanceof Error ? e.message : String(e)
  }
}

async function save(): Promise<void> {
  const { data } = await saveXlsx(local.value)
  const a = document.createElement('a')
  a.href = URL.createObjectURL(new Blob([data as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
  a.download = fileName.value
  a.click()
  URL.revokeObjectURL(a.href)
}
onBeforeUnmount(() => void local.value.close())
const ms = (n: number) => `${Math.round(n)} мс`
</script>

<template>
  <div class="page">
    <header class="top">
      <strong>@xlcalc/vue/local</strong>
      <span class="state" data-calc>{{ calc === 'pool' ? 'Web Worker' : 'тот же поток' }}</span>
      <span class="state" data-timings>
        книга {{ ms(timings.built) }} · расчёт {{ timings.ready ? ms(timings.ready) : '…' }}<template v-if="timings.edit"> · правка {{ ms(timings.edit) }}, ячеек {{ timings.changes }}</template>
      </span>
      <span v-if="failed" class="error">{{ failed }}</span>
      <label class="file">Открыть xlsx <input type="file" accept=".xlsx" data-open @change="openFile" /></label>
      <button type="button" @click="save">Сохранить xlsx</button>
      <span class="state" data-file>{{ fileName }}</span>
      <span class="state">{{ state ? `${state.sheet}!${state.ref}${state.formula ? ' =' + state.formula : ''}` : '' }}</span>
    </header>
    <div v-if="lost.length" class="lost" data-lost>
      Из файла не перенесено:
      <span v-for="g in lost" :key="g.code">{{ g.message }}<template v-if="g.count > 1"> (и ещё {{ g.count - 1 }})</template><template v-if="g.sheets.length">, листы {{ g.sheets.join(', ') }}</template>. </span>
    </div>
    <SheetView :key="fileName" class="sheet" :source="source" :messages="ruMessages" :locale="local.locale" :focus="focus" storage-key="xs-local" @state="state = $event" />
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
.file input {
  max-width: 180px;
}
.lost {
  padding: 4px 10px;
  background: #fff7e0;
  border-bottom: 1px solid #f0d58a;
  color: #6b4e00;
}
.sheet {
  flex: 1 1 auto;
  min-height: 0;
}
@media (max-width: 600px) {
  .top .state:last-of-type,
  .top button,
  .top .file {
    display: none;
  }
}
</style>
