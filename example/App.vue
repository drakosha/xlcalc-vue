<script setup lang="ts">
// The example page: the sheet over the "all formatting" book, a language switch and the next edit's outcome for trying refusals.
import { ref, shallowRef } from 'vue'
import { SheetView, enMessages, memorySource, ruMessages, type ListSort, type SheetLinkEvent, type SheetViewState } from '../src'
import { allFormattingBook } from './book'
import { LIST_SHEET, sortList } from './list-book'

const params = new URLSearchParams(location.search)
const lang = ref(params.get('lang') === 'en' ? 'en' : 'ru')
const source = memorySource(allFormattingBook(Number(params.get('rows') ?? 34_000)), { delay: Number(params.get('delay') ?? 150) })
const focus = shallowRef(params.get('sheet') ? { ref: params.get('sheet')! } : null)
const state = ref<SheetViewState | null>(null)
const lastLink = ref('')
const next = ref('')

/** The next edit's refusal, to see how the sheet answers one; `?fail=conflict` sets it on load. */
function failNext(kind: string): void {
  if (kind === 'conflict') source.failNext({ ok: false, reason: 'conflict', message: '' })
  else if (kind === 'invalid') source.failNext({ ok: false, reason: 'invalid', message: lang.value === 'ru' ? 'Формула не разобрана: лишняя скобка' : 'The formula does not parse: an extra bracket', at: 2 })
  else if (kind === 'refused') source.failNext({ ok: false, reason: 'refused', message: lang.value === 'ru' ? 'Ячейку менять нельзя: область скрипта' : 'This cell belongs to the script' })
  next.value = ''
}
failNext(params.get('fail') ?? '')
function onLink(link: SheetLinkEvent): void {
  lastLink.value = `${link.sheet}!${link.cell} → ${link.target ?? link.href ?? ''}`
}
/** A book being recomputed for a few seconds: the sheet waits with a stopwatch; `?recompute=4000` plays it on load. */
function recompute(ms = 4000): void {
  source.recompute(ms)
  source.emit({ kind: 'reload' })
}
if (params.get('recompute')) source.recompute(Number(params.get('recompute')))
/** Another tab's edit: the sheet gets it as a `changed` event. */
function remoteEdit(): void {
  source.setCells('Формулы', { A2: String(Math.round(Math.random() * 100)) })
}
/** The list sheet: the page sorts it as a host sorts its list, and shows what a drill-down or a row opened. */
const listSort = ref<ListSort | null>(null)
function onSort(sort: ListSort): void {
  listSort.value = sort
  sortList(source, sort)
}
const onDrill = (at: { sheet: string; cell: string }) => (lastLink.value = `cellDrill ${at.sheet}!${at.cell}`)
const onRowOpen = (at: { sheet: string; row: number }) => (lastLink.value = `rowOpen ${at.sheet} ${at.row}`)
const cellUrl = (place: string) => `${location.origin}${location.pathname}?sheet=${encodeURIComponent(place)}`
</script>

<template>
  <div class="page">
    <header class="top">
      <strong>@xlcalc/vue</strong>
      <select v-model="lang" aria-label="language">
        <option value="ru">Русский</option>
        <option value="en">English</option>
      </select>
      <label>
        {{ lang === 'ru' ? 'Следующая правка' : 'Next edit' }}
        <select v-model="next" @change="failNext(next)">
          <option value="">ok</option>
          <option value="conflict">conflict</option>
          <option value="invalid">invalid</option>
          <option value="refused">refused</option>
        </select>
      </label>
      <button type="button" @click="recompute()">{{ lang === 'ru' ? 'Пересчёт 4 с' : 'Recompute 4 s' }}</button>
      <button type="button" @click="remoteEdit">{{ lang === 'ru' ? 'Правка из другой вкладки' : 'Edit from another tab' }}</button>
      <span class="state" data-state>{{ state ? `${state.sheet}!${state.ref}${state.formula ? ' =' + state.formula : ''}` : '' }}</span>
      <span v-if="lastLink" class="state">{{ lastLink }}</span>
    </header>
    <SheetView
      class="sheet"
      :source="source"
      :messages="lang === 'ru' ? ruMessages : enMessages"
      locale="ru-RU"
      storage-key="xs-example"
      :focus="focus"
      :cell-url="cellUrl"
      :sort="listSort"
      :open-rows="state?.sheet === LIST_SHEET"
      @state="state = $event"
      @open-link="onLink"
      @sort="onSort"
      @cell-drill="onDrill"
      @row-open="onRowOpen"
    />
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
.sheet {
  flex: 1 1 auto;
  min-height: 0;
}
@media (max-width: 600px) {
  .top .state,
  .top label,
  .top button {
    display: none;
  }
}
</style>
