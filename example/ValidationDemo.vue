<script setup lang="ts">
// A page for the components of data validation until SheetView mounts them itself: the sheet over the "all formatting"
// book, the dropdown and the input message teleported into its grid at the active cell, the alert into the sheet. They
// are wired from outside: the active cell is read from the grid, Alt+Down is caught before the grid's keys, an item goes
// in as another tab's edit, and a refused edit's alert is answered inside the source (Yes and OK send it with `accept`).
import { onBeforeUnmount, onMounted, ref, shallowRef } from 'vue'
import { SheetView, enMessages, memorySource, ruMessages, type SheetViewState } from '../src'
import ValidationAlert from '../src/components/ValidationAlert.vue'
import ValidationDropdown from '../src/components/ValidationDropdown.vue'
import ValidationPrompt from '../src/components/ValidationPrompt.vue'
import { useValidation, useValidationAlert, type AlertAnswer, type ValidationRect } from '../src/core/validation'
import type { EditOutcome, SheetSource } from '../src/source'
import { allFormattingBook } from './book'

const params = new URLSearchParams(location.search)
const messages = params.get('lang') === 'en' ? enMessages : ruMessages
const focus = params.get('sheet') ? { ref: params.get('sheet')! } : null
const book = memorySource(allFormattingBook(100), { delay: 60 })
const alerts = useValidationAlert()
/** The book with Excel's alert answered before the sheet gets the outcome. */
const source: SheetSource = {
  ...book,
  async edit(entries, opts): Promise<EditOutcome> {
    const got = await book.edit!(entries, opts)
    if (got.ok || !got.alert) return got
    const answer = await alerts.ask(got.alert)
    return answer === 'accept' ? book.edit!(entries, { accept: true }) : { ok: true, changes: [] }
  },
}

const host = ref<HTMLElement | null>(null)
const scroller = shallowRef<HTMLElement | null>(null)
const frame = shallowRef<HTMLElement | null>(null)
const state = ref<SheetViewState | null>(null)
const active = ref<{ row: number; col: number } | null>(null)
const rect = ref<ValidationRect | null>(null)
const view = ref<ValidationRect | null>(null)
const touch = matchMedia('(pointer: coarse)').matches || innerWidth < 600

/** The active cell and its rectangle in the grid's content, read from the drawn grid. */
function measure(): void {
  const sc = scroller.value
  const td = sc?.querySelector<HTMLElement>('.xs-cell.xs-active')
  if (!sc || !td) return void (rect.value = null)
  const [row, col] = td.dataset.cell!.split(':').map(Number) as [number, number]
  if (active.value?.row !== row || active.value?.col !== col) active.value = { row, col }
  const s = sc.getBoundingClientRect()
  const r = td.getBoundingClientRect()
  rect.value = { top: r.top - s.top + sc.scrollTop, left: r.left - s.left + sc.scrollLeft, width: r.width, height: r.height }
  view.value = { top: sc.scrollTop, left: sc.scrollLeft, width: sc.clientWidth, height: sc.clientHeight }
}
const cellText = () => scroller.value?.querySelector('.xs-cell.xs-active')?.textContent?.trim() ?? ''

const val = useValidation({
  source: () => source,
  sheet: () => state.value?.sheet ?? '',
  active: () => active.value,
  text: cellText,
  pick: (text, at) => void book.setCells(at.sheet, { [at.cell]: text }),
})
function onState(next: SheetViewState | null): void {
  state.value = next
  requestAnimationFrame(measure)
}
/** Alt+Down and the open list's keys before the grid's. */
function onKey(e: KeyboardEvent): void {
  if (e.target === scroller.value && val.onKey(e)) e.stopPropagation()
}
/** The alert answered: the grid takes its keys back. */
function answer(a: AlertAnswer): void {
  alerts.answer(a)
  scroller.value?.focus({ preventScroll: true })
}
let timer: ReturnType<typeof setInterval> | null = null
onMounted(() => {
  // the grid mounts after its sheets come: wait for its scroller
  timer = setInterval(() => {
    const sc = host.value?.querySelector<HTMLElement>('.xs-scroller')
    if (!sc) return
    scroller.value = sc
    frame.value = host.value!.querySelector<HTMLElement>('.xs-sheet')
    sc.addEventListener('scroll', measure, { passive: true })
    measure()
    clearInterval(timer!)
  }, 50)
})
onBeforeUnmount(() => timer && clearInterval(timer))
</script>

<template>
  <div ref="host" class="demo" @keydown.capture="onKey">
    <SheetView class="sheet" :source="source" :messages="messages" locale="ru-RU" :focus="focus" @state="onState" />
    <Teleport v-if="scroller && rect" :to="scroller">
      <ValidationDropdown
        v-if="val.list.value"
        :rect="rect"
        :items="val.list.value"
        :open="val.open.value"
        :active="val.index.value"
        :view="view ?? undefined"
        :touch="touch"
        :messages="messages"
        @toggle="val.toggle()"
        @pick="(i) => val.choose(i)"
        @hover="(i) => (val.index.value = i)"
        @close="val.close()"
      />
      <ValidationPrompt v-if="val.prompt.value" :rect="rect" :prompt="val.prompt.value" :view="view ?? undefined" />
    </Teleport>
    <Teleport v-if="frame" :to="frame">
      <ValidationAlert :alert="alerts.alert.value" :messages="messages" @answer="answer" />
    </Teleport>
  </div>
</template>

<style>
.demo {
  height: 100%;
}
.sheet {
  height: 100%;
}
</style>
