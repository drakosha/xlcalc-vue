<script setup lang="ts">
// An app like Excel in the browser over the local source: a new book in a chosen language, an xlsx file opened from the
// picker or dropped on the page (in the interface's language when the file names none, on its first sheet with data),
// the formatting toolbar over the sheet with Excel's Format Cells dialog, Save as under the book's name, what the file
// lost, a status bar (reading, computing, ready) and the book's settings. The book computes in a Web Worker; nothing is sent.
import { computed, onBeforeUnmount, ref, shallowRef, watchEffect } from 'vue'
import type { BookInput } from 'xlcalc'
import { quoteSheetIfNeeded } from 'xlcalc/text'
import { FormatDialog, SheetToolbar, SheetView, enMessages, ruMessages, type SheetFormatPatch, type SheetViewApi, type SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import SpaDialog from './SpaDialog.vue'
import SpaLosses from './SpaLosses.vue'
import { download, emptyBook, timed } from './spa-book'
import { IconNew, IconOpen, IconSave, IconSettings } from './spa-icons'
import { bookLocaleOf, firstShownSheet, localeOfLang, lossesOf, type LossKind } from './spa-open'
import { readInThread, writeInThread } from './spa-files'
import { BOOK_LOCALES, pickLang, TEXTS, type BookLocale, type SpaLang } from './spa-texts'

const params = new URLSearchParams(location.search)
const lang = ref<SpaLang>(pickLang(params.get('lang')))
const t = computed(() => TEXTS[lang.value])
const messages = computed(() => (lang.value === 'ru' ? ruMessages : enMessages))
watchEffect(() => (document.documentElement.lang = lang.value))

/** The open book: its source, a key that mounts the sheet anew for each book, what its file lost and the sheet it opens on. */
interface Opened {
  source: LocalSource
  key: number
  lost: LossKind[]
  focus: { ref: string } | null
}
const book = shallowRef<Opened | null>(null)
const name = ref('')
const state = ref<SheetViewState | null>(null)
/** The sheet the toolbar formats, undoes and redoes through, as the sheet's own keys do. */
const sheet = shallowRef<SheetViewApi | null>(null)
const lostShown = ref(true)
let keys = 0
let created = 0

// ── the status bar: reading a file, computing its book, ready, or calculating while a change is out ──
const reading = ref(false)
const opening = ref(false)
const saving = ref(false)
const out = ref(0)
const lastRead = ref(0)
const lastOpen = ref(0)
const lastChange = ref(0)
const lastSave = ref(0)
const note = ref('')
let noteTimer: ReturnType<typeof setTimeout> | null = null
function say(text: string): void {
  if (noteTimer) clearTimeout(noteTimer)
  note.value = text
  noteTimer = setTimeout(() => (note.value = ''), 8000)
}
const busy = computed(() => reading.value || opening.value || out.value > 0)
const status = computed(() => {
  const x = t.value
  return saving.value ? x.saving : reading.value ? x.reading : opening.value ? x.computing : out.value > 0 ? x.busy : x.ready
})
const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e))

/** Open a book on a sheet: a new source with its changes timed, the old one closed; `since` is when the opening began. */
function start(input: BookInput, bookName: string, lost: LossKind[] = [], since = performance.now(), sheet?: string): void {
  const previous = book.value?.source
  const source = timed(
    localSource(input),
    () => out.value++,
    (ms) => {
      out.value--
      lastChange.value = ms
    },
  )
  book.value = { source, key: ++keys, lost, focus: sheet === undefined ? null : { ref: quoteSheetIfNeeded(sheet) } }
  name.value = bookName
  lostShown.value = true
  state.value = null
  lastChange.value = 0
  opening.value = true
  const mine = () => book.value?.source === source
  source.ready
    .then(
      () => {
        if (!mine()) return
        lastOpen.value = performance.now() - since
      },
      (e: unknown) => mine() && say(t.value.openFailed(errorText(e))),
    )
    .finally(() => mine() && (opening.value = false))
  void previous?.close()
}
start(emptyBook(localeOfLang(lang.value)), t.value.defaultName(++created))

// ── New ──
const newOpen = ref(false)
const newLocale = ref<BookLocale>(localeOfLang(lang.value))
function create(): void {
  newOpen.value = false
  start(emptyBook(newLocale.value), t.value.defaultName(++created))
}

// ── Open: the picker or a file dropped on the page ──
const picker = ref<HTMLInputElement | null>(null)
async function openFile(file: File): Promise<void> {
  const since = performance.now()
  reading.value = true
  try {
    // read in a thread of its own; what the file lost, in the interface's language
    const read = await readInThread(file, localeOfLang(lang.value))
    lastRead.value = performance.now() - since
    const locale = bookLocaleOf(read.book.locale, lang.value)
    start({ ...read.book, locale }, file.name.replace(/\.xlsx$/i, ''), lossesOf(read.issues), since, firstShownSheet(read.book))
  } catch (e) {
    say(t.value.openFailed(errorText(e)))
  } finally {
    reading.value = false
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

// ── Save as: the book with its values (saveXlsx of ./local, the writing in the file thread), downloaded under its name ──
async function saveAs(): Promise<void> {
  const b = book.value
  if (!b) return
  saving.value = true
  const since = performance.now()
  try {
    const data = await writeInThread(await b.source.snapshot())
    download(data, `${name.value.trim().replace(/\.xlsx$/i, '') || t.value.defaultName(created)}.xlsx`)
    lastSave.value = performance.now() - since
  } catch (e) {
    say(t.value.saveFailed(errorText(e)))
  } finally {
    saving.value = false
  }
}

// ── the book's settings: its language (the book opens again in it, on the same sheet) and the interface's ──
const settingsOpen = ref(false)
const draftLocale = ref<BookLocale>('en-US')
const draftLang = ref<SpaLang>(lang.value)
function openSettings(): void {
  draftLocale.value = (book.value?.source.locale as BookLocale | undefined) ?? localeOfLang(lang.value)
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
    start({ ...input, locale: draftLocale.value }, name.value, b.lost, performance.now(), state.value?.sheet)
  } catch (e) {
    say(t.value.openFailed(errorText(e)))
  }
}

// ── the grid's focus back after the toolbar's number format list or the Format Cells dialog took it ──
function refocus(): void {
  if (!window.matchMedia?.('(pointer: coarse)').matches) document.querySelector<HTMLElement>('.xs-scroller')?.focus({ preventScroll: true })
}

// ── Format Cells: the toolbar's button opens it, OK formats the selection through the sheet, one step of undo ──
const formatOpen = ref(false)
function formatCells(patches: SheetFormatPatch[]): void {
  void sheet.value?.format(patches)
}
function closeFormat(): void {
  formatOpen.value = false
  refocus()
}

// The browser check reads the timings and the book through this.
;(window as unknown as { __spa: unknown }).__spa = {
  timings: () => ({ read: lastRead.value, open: lastOpen.value, change: lastChange.value, save: lastSave.value }),
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
        <input v-model="name" class="spa-name-input" :size="Math.min(60, Math.max(12, name.length + 1))" data-spa-name :aria-label="t.bookName" spellcheck="false" autocomplete="off" />
        <span class="spa-ext">.xlsx</span>
      </label>
      <button type="button" class="spa-btn" data-spa="settings" :title="t.settings" :aria-label="t.settings" @click="openSettings">
        <IconSettings />
        <span class="spa-label">{{ t.settings }}</span>
      </button>
    </header>
    <SpaLosses v-if="book && book.lost.length && lostShown" :key="book.key" :kinds="book.lost" :t="t" @close="lostShown = false" />
    <SheetToolbar
      v-if="book"
      :source="book.source"
      :state="state"
      :messages="messages"
      :locale="book.source.locale"
      :sheet="sheet"
      @done="refocus"
      @format-cells="formatOpen = true"
    />
    <FormatDialog :open="formatOpen" :state="state" :messages="messages" :locale="book?.source.locale" @apply="formatCells" @close="closeFormat" />
    <SheetView
      v-if="book"
      :key="book.key"
      :ref="(el) => (sheet = el as SheetViewApi | null)"
      class="spa-sheet"
      :source="book.source"
      :messages="messages"
      :locale="book.source.locale"
      :focus="book.focus"
      @state="state = $event"
    />
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
