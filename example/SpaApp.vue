<script setup lang="ts">
// An app like Excel in the browser over the local source: a new book in a chosen language, an xlsx file opened from the
// picker or dropped on the page, the formatting toolbar over the sheet, Save as under the book's name, what the file
// lost, a status bar (ready or calculating) and the book's settings. The book computes in a Web Worker; nothing is sent.
import { computed, onBeforeUnmount, ref, shallowRef, watchEffect } from 'vue'
import type { BookInput } from 'xlcalc'
import { quoteSheetIfNeeded } from 'xlcalc/a1'
import { SheetToolbar, SheetView, enMessages, ruMessages, type SheetViewState, type ToolbarNumberFormat } from '../src'
import { groupIssues, localSource, openXlsx, saveXlsx, type IssueGroup, type LocalSource } from '../src/local'
import SpaDialog from './SpaDialog.vue'
import { download, emptyBook, numberFormatsIn, timed } from './spa-book'
import { IconClose, IconNew, IconOpen, IconSave, IconSettings } from './spa-icons'
import { BOOK_LOCALES, pickLang, TEXTS, type BookLocale, type SpaLang } from './spa-texts'

const params = new URLSearchParams(location.search)
const lang = ref<SpaLang>(pickLang(params.get('lang')))
const t = computed(() => TEXTS[lang.value])
const messages = computed(() => (lang.value === 'ru' ? ruMessages : enMessages))
const langLocale = (l: SpaLang): BookLocale => (l === 'ru' ? 'ru-RU' : 'en-US')
watchEffect(() => (document.documentElement.lang = lang.value))

/** The open book: its source, a key that mounts the sheet anew for each book, and what its file lost. */
interface Opened {
  source: LocalSource
  key: number
  lost: IssueGroup[]
}
const book = shallowRef<Opened | null>(null)
const name = ref('')
const formats = shallowRef<ToolbarNumberFormat[]>([])
const state = ref<SheetViewState | null>(null)
const focus = shallowRef<{ ref: string } | null>(null)
const lostShown = ref(true)
let keys = 0
let created = 0

// ── the status bar: ready, or calculating while a book opens or a change is out ──
const opening = ref(false)
const saving = ref(false)
const out = ref(0)
const lastOpen = ref(0)
const lastChange = ref(0)
const note = ref('')
let noteTimer: ReturnType<typeof setTimeout> | null = null
function say(text: string): void {
  if (noteTimer) clearTimeout(noteTimer)
  note.value = text
  noteTimer = setTimeout(() => (note.value = ''), 8000)
}
const busy = computed(() => opening.value || out.value > 0)
const status = computed(() => (saving.value ? t.value.saving : busy.value ? t.value.busy : t.value.ready))
const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e))

/** Open a book: a new source with its changes timed, the old one closed; `since` is when the opening began. */
function start(input: BookInput, bookName: string, lost: IssueGroup[] = [], since = performance.now()): void {
  const previous = book.value?.source
  const source = timed(
    localSource(input),
    () => out.value++,
    (ms) => {
      out.value--
      lastChange.value = ms
    },
  )
  book.value = { source, key: ++keys, lost }
  name.value = bookName
  lostShown.value = true
  state.value = null
  focus.value = null
  formats.value = []
  lastChange.value = 0
  opening.value = true
  const mine = () => book.value?.source === source
  source.ready
    .then(
      () => {
        if (!mine()) return
        lastOpen.value = performance.now() - since
        formats.value = numberFormatsIn(source.locale)
      },
      (e: unknown) => mine() && say(t.value.openFailed(errorText(e))),
    )
    .finally(() => mine() && (opening.value = false))
  void previous?.close()
}
start(emptyBook(langLocale(lang.value)), t.value.defaultName(++created))

// ── New ──
const newOpen = ref(false)
const newLocale = ref<BookLocale>(langLocale(lang.value))
function create(): void {
  newOpen.value = false
  start(emptyBook(newLocale.value), t.value.defaultName(++created))
}

// ── Open: the picker or a file dropped on the page ──
const picker = ref<HTMLInputElement | null>(null)
async function openFile(file: File): Promise<void> {
  const since = performance.now()
  opening.value = true
  try {
    const read = await openXlsx(file)
    start(read.book, file.name.replace(/\.xlsx$/i, ''), groupIssues(read.issues), since)
  } catch (e) {
    opening.value = false
    say(t.value.openFailed(errorText(e)))
  }
}
function onPick(e: Event): void {
  const input = e.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (file) void openFile(file)
}
const dragging = ref(false)
function onDragOver(e: DragEvent): void {
  if (!e.dataTransfer?.types.includes('Files')) return
  e.preventDefault()
  dragging.value = true
}
function onDragLeave(e: DragEvent): void {
  if (!e.relatedTarget) dragging.value = false
}
function onDrop(e: DragEvent): void {
  if (!e.dataTransfer?.types.includes('Files')) return
  e.preventDefault()
  dragging.value = false
  const file = e.dataTransfer.files[0]
  if (file) void openFile(file)
}

// ── Save as: the book with its values, downloaded under its name ──
async function saveAs(): Promise<void> {
  const b = book.value
  if (!b) return
  saving.value = true
  try {
    const { data } = await saveXlsx(b.source)
    download(data, `${name.value.trim().replace(/\.xlsx$/i, '') || t.value.defaultName(created)}.xlsx`)
  } catch (e) {
    say(t.value.saveFailed(errorText(e)))
  } finally {
    saving.value = false
  }
}

// ── the book's settings: its language (the book opens again in it) and the interface's ──
const settingsOpen = ref(false)
const draftLocale = ref<BookLocale>('en-US')
const draftLang = ref<SpaLang>(lang.value)
function openSettings(): void {
  draftLocale.value = (book.value?.source.locale as BookLocale | undefined) ?? langLocale(lang.value)
  draftLang.value = lang.value
  settingsOpen.value = true
}
async function applySettings(): Promise<void> {
  settingsOpen.value = false
  lang.value = draftLang.value
  const b = book.value
  if (!b || draftLocale.value === b.source.locale) return
  try {
    const input = await b.source.book()
    start({ ...input, locale: draftLocale.value }, name.value, b.lost)
  } catch (e) {
    say(t.value.openFailed(errorText(e)))
  }
}

// ── the toolbar's outcomes: the selection after undo, refusals, the grid's focus back ──
function onSelect(place: { sheet: string; range: string }): void {
  focus.value = { ref: `${quoteSheetIfNeeded(place.sheet)}!${place.range}` }
}
function refocus(): void {
  if (!window.matchMedia?.('(pointer: coarse)').matches) document.querySelector<HTMLElement>('.xs-scroller')?.focus({ preventScroll: true })
}

// The browser check reads the timings and the book through this.
;(window as unknown as { __spa: unknown }).__spa = {
  timings: () => ({ open: lastOpen.value, change: lastChange.value }),
  book: () => book.value?.source.book(),
  status: () => status.value,
}
onBeforeUnmount(() => void book.value?.source.close())
</script>

<template>
  <div :class="['spa', { 'spa-dragging': dragging }]" @dragover="onDragOver" @dragleave="onDragLeave" @drop="onDrop">
    <header class="spa-top">
      <span class="spa-mark" aria-hidden="true">xlcalc</span>
      <button type="button" class="spa-btn" data-spa="new" :title="t.newBook" :aria-label="t.newBook" @click="newOpen = true">
        <IconNew />
        <span class="spa-label">{{ t.newBook }}</span>
      </button>
      <button type="button" class="spa-btn" data-spa="open" :title="t.open" :aria-label="t.open" @click="picker?.click()">
        <IconOpen />
        <span class="spa-label">{{ t.open }}</span>
      </button>
      <input ref="picker" class="spa-file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" tabindex="-1" data-spa-file @change="onPick" />
      <button type="button" class="spa-btn" data-spa="save" :title="t.saveAs" :aria-label="t.saveAs" :disabled="!book || saving" @click="saveAs">
        <IconSave />
        <span class="spa-label">{{ t.saveAs }}</span>
      </button>
      <label class="spa-name" :title="t.bookName">
        <input v-model="name" class="spa-name-input" data-spa-name :aria-label="t.bookName" spellcheck="false" autocomplete="off" />
        <span class="spa-ext">.xlsx</span>
      </label>
      <button type="button" class="spa-btn" data-spa="settings" :title="t.settings" :aria-label="t.settings" @click="openSettings">
        <IconSettings />
        <span class="spa-label">{{ t.settings }}</span>
      </button>
    </header>
    <div v-if="book && book.lost.length && lostShown" class="spa-lost" role="status" data-spa-lost>
      <p>
        {{ t.lost }}
        <span v-for="g in book.lost" :key="g.code">{{ g.message }}{{ g.count > 1 ? t.lostMore(g.count - 1) : '' }}{{ g.sheets.length ? t.lostSheets(g.sheets) : '' }}. </span>
      </p>
      <button type="button" class="spa-btn spa-icon-btn" :title="t.close" :aria-label="t.close" @click="lostShown = false"><IconClose /></button>
    </div>
    <SheetToolbar
      v-if="book"
      :source="book.source"
      :state="state"
      :messages="messages"
      :number-formats="formats"
      @select="onSelect"
      @error="say"
      @done="refocus"
    />
    <SheetView v-if="book" :key="book.key" class="spa-sheet" :source="book.source" :messages="messages" :locale="book.source.locale" :focus="focus" @state="state = $event" />
    <footer class="spa-status">
      <span :class="['spa-state', { 'spa-busy': busy || saving }]" role="status" data-spa-status>{{ status }}</span>
      <span v-if="note" class="spa-note" data-spa-note>{{ note }}</span>
      <span class="spa-meta">
        {{ book?.source.locale }}<template v-if="lastOpen"> · {{ t.opened(lastOpen) }}</template><template v-if="lastChange"> · {{ t.edited(lastChange) }}</template>
      </span>
    </footer>
    <div v-if="dragging" class="spa-drop" aria-hidden="true">{{ t.dropHere }}</div>

    <SpaDialog v-if="newOpen" :title="t.newTitle" @close="newOpen = false">
      <fieldset class="spa-field">
        <legend>{{ t.bookLocale }}</legend>
        <label v-for="l in BOOK_LOCALES" :key="l" class="spa-choice">
          <input v-model="newLocale" type="radio" name="spa-new-locale" :value="l" :data-spa-locale="l" />
          {{ t.localeName[l] }}
        </label>
      </fieldset>
      <template #foot>
        <button type="button" class="spa-btn" @click="newOpen = false">{{ t.cancel }}</button>
        <button type="submit" class="spa-btn spa-primary" data-spa="create" @click="create">{{ t.create }}</button>
      </template>
    </SpaDialog>

    <SpaDialog v-if="settingsOpen" :title="t.settings" @close="settingsOpen = false">
      <label class="spa-field">
        <span>{{ t.bookLocale }}</span>
        <select v-model="draftLocale" class="spa-select" data-spa-book-locale>
          <option v-for="l in BOOK_LOCALES" :key="l" :value="l">{{ t.localeName[l] }}</option>
        </select>
        <small class="spa-hint">{{ t.bookLocaleHint }}</small>
      </label>
      <label class="spa-field">
        <span>{{ t.uiLang }}</span>
        <select v-model="draftLang" class="spa-select" data-spa-ui-lang>
          <option v-for="l in ['ru', 'en'] as SpaLang[]" :key="l" :value="l">{{ t.langName[l] }}</option>
        </select>
      </label>
      <template #foot>
        <button type="button" class="spa-btn" @click="settingsOpen = false">{{ t.cancel }}</button>
        <button type="submit" class="spa-btn spa-primary" data-spa="apply" @click="applySettings">{{ t.apply }}</button>
      </template>
    </SpaDialog>
  </div>
</template>

<style src="./spa.css"></style>
