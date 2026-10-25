<script setup lang="ts">
// Excel's Symbol dialog (Insert > Symbol): the Symbols tab (the Subset box, the character grid, the recently used symbols,
// the Character code box) and the Special Characters tab, in the active cell's font; Insert puts the character into the
// edit, which stays open while the dialog does, and the dialog stays for more until Close.
import { computed, onBeforeUnmount, onMounted, ref, useId, watch } from 'vue'
import { fontCss } from '../core/cell-style'
import { listStep, stepIndex } from '../core/list-keys'
import { codeHex, DEFAULT_RECENT_SYMBOLS, parseCode, RECENT_SYMBOLS, SYMBOL_SPECIALS, SYMBOL_SUBSETS, subsetOf, symbolCodes } from '../symbols'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { keepActiveInView } from './popup'
import '../controls.css'

const ctx = useSheetContext()
const m = ctx.messages
const id = useId()

/** The recently used symbols, newest first, kept in the sheet's storage; Office's row until something is inserted. */
const RECENT = 'recentSymbols'
const stored = ctx.memory.recall<number[]>(RECENT)?.filter((c) => Number.isInteger(c) && parseCode(codeHex(c)) === c) ?? null
const recent = ref<number[]>(stored ?? [...DEFAULT_RECENT_SYMBOLS])

const codes = symbolCodes()
const indexOf = new Map(codes.map((c, i) => [c, i]))
const tab = ref<'symbols' | 'special'>('symbols')
/** The selected character: the last one inserted, otherwise the grid's first. */
const code = ref(stored?.[0] ?? codes[0]!)
const codeText = ref(codeHex(code.value))
const special = ref(0)
/** Cancel becomes Close once something has been inserted, as in Excel. */
const inserted = ref(false)
/** The characters are shown in the active cell's font, Excel's "(normal text)". */
const font = computed(() => fontCss(ctx.viewState.value?.style, undefined).fontFamily)

// ── the grid: as many columns as fit (16 at most, as Excel's), only the rows in view drawn ──
const grid = ref<HTMLElement | null>(null)
const box = ref({ width: 0, height: 0 })
const scrollTop = ref(0)
const cellPx = computed(() => (ctx.v.touch.value ? 44 : 30))
const cols = computed(() => (box.value.width ? Math.min(16, Math.max(6, Math.floor(box.value.width / cellPx.value))) : 16))
const rowCount = computed(() => Math.ceil(codes.length / cols.value))
const viewRows = computed(() => Math.max(1, Math.floor((box.value.height || 8 * cellPx.value) / cellPx.value)))
const rows = computed(() => {
  const first = Math.max(0, Math.floor(scrollTop.value / cellPx.value) - 2)
  const last = Math.min(rowCount.value, first + viewRows.value + 5)
  const out: { row: number; codes: number[] }[] = []
  for (let r = first; r < last; r++) out.push({ row: r, codes: codes.slice(r * cols.value, (r + 1) * cols.value) })
  return out
})
const at = computed(() => indexOf.get(code.value) ?? -1)
const subset = computed(() => subsetOf(code.value) ?? '')

function measure(): void {
  const el = grid.value
  if (el) box.value = { width: el.clientWidth, height: el.clientHeight }
}
/** Scrolls the grid so the selected character's row is in view. */
function reveal(): void {
  const el = grid.value
  if (!el || at.value < 0) return
  const top = Math.floor(at.value / cols.value) * cellPx.value
  const height = el.clientHeight || viewRows.value * cellPx.value
  if (top < el.scrollTop) el.scrollTop = top
  else if (top + cellPx.value > el.scrollTop + height) el.scrollTop = top + cellPx.value - height
  scrollTop.value = el.scrollTop
}
/** Selects a character; the Character code box shows its code unless the code was typed there. */
function select(c: number, typed = false): void {
  code.value = c
  if (!typed) codeText.value = codeHex(c)
  reveal()
}
const observer = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(measure) : null
onMounted(() => {
  measure()
  if (grid.value) observer?.observe(grid.value)
  reveal()
})
onBeforeUnmount(() => observer?.disconnect())
watch(cols, reveal, { flush: 'post' })

/** The Subset box: the subset's first character, at the top of the grid. */
function toSubset(e: Event): void {
  const s = SYMBOL_SUBSETS.find((x) => x.key === (e.target as HTMLSelectElement).value)
  const first = s && codes.find((c) => c >= s.from && c <= s.to)
  if (first === undefined || !grid.value) return
  select(first)
  grid.value.scrollTop = Math.floor(at.value / cols.value) * cellPx.value
  scrollTop.value = grid.value.scrollTop
}
function onCodeInput(e: Event): void {
  codeText.value = (e.target as HTMLInputElement).value
  const c = parseCode(codeText.value)
  if (c !== null) select(c, true)
}
/** Arrows by one character and by a row, Home and End to the row's ends (with Ctrl, the grid's), the page keys by the rows in view. */
function onGridKey(e: KeyboardEvent): void {
  const i = Math.max(0, at.value)
  const col = i % cols.value
  const by: Record<string, number> = {
    ArrowRight: 1,
    ArrowLeft: -1,
    ArrowDown: cols.value,
    ArrowUp: -cols.value,
    PageDown: viewRows.value * cols.value,
    PageUp: -viewRows.value * cols.value,
    Home: e.ctrlKey || e.metaKey ? -i : -col,
    End: e.ctrlKey || e.metaKey ? codes.length : cols.value - 1 - col,
  }
  const step = by[e.key]
  if (step === undefined) return
  e.preventDefault()
  select(codes[stepIndex(i, step, codes.length)]!)
}
const specials = ref<HTMLElement | null>(null)
keepActiveInView(specials, special, () => special.value)
function onSpecialKey(e: KeyboardEvent): void {
  const step = listStep(e.key, SYMBOL_SPECIALS.length, 8)
  if (step === null) return
  e.preventDefault()
  special.value = stepIndex(special.value, step, SYMBOL_SPECIALS.length)
}

/** The tab's character: the grid's, or the Special Characters tab's. */
const chosen = computed(() => (tab.value === 'special' ? SYMBOL_SPECIALS[special.value]!.code : code.value))
function insert(): void {
  const c = chosen.value
  if (!ctx.ed.insertText(String.fromCodePoint(c))) return
  inserted.value = true
  recent.value = [c, ...recent.value.filter((x) => x !== c)].slice(0, RECENT_SYMBOLS)
  ctx.memory.remember(RECENT, recent.value)
}
/** Enter means Insert, except on a button or a list box of its own. */
function onKey(e: KeyboardEvent): void {
  const t = e.target as HTMLElement
  if (e.key !== 'Enter' || e.isComposing || t.closest('button, select')) return
  e.preventDefault()
  insert()
}
</script>

<template>
  <DialogShell class="xs-sym" full focus="[data-xs-sym-grid]" :aria-label="m.symbolTitle" data-xs-keep-edit data-xs-symbol @keydown="onKey" @close="ctx.closeSymbol()">
    <p class="xs-modal-title">{{ m.symbolTitle }}</p>
    <div class="xs-sym-tabs" role="tablist">
      <button v-for="t in ['symbols', 'special'] as const" :key="t" type="button" class="xs-sym-tab" role="tab" :aria-selected="tab === t" :data-xs-sym-tab="t" @click="tab = t">
        {{ m.symbolTab(t) }}
      </button>
    </div>

    <div v-show="tab === 'symbols'" class="xs-sym-panel" role="tabpanel" :aria-label="m.symbolTab('symbols')">
      <label class="xs-sym-field">
        <span>{{ m.symbolSubsetLabel }}</span>
        <select class="xs-field xs-field-l" :value="subset" data-xs-sym-subset @change="toSubset">
          <option v-for="s in SYMBOL_SUBSETS" :key="s.key" :value="s.key">{{ m.symbolSubset(s.key) }}</option>
        </select>
      </label>
      <div
        ref="grid"
        class="xs-sym-grid"
        role="listbox"
        tabindex="0"
        :aria-label="m.symbolTitle"
        :aria-activedescendant="at >= 0 ? `${id}-c${code}` : undefined"
        :style="{ fontFamily: font }"
        data-xs-sym-grid
        @keydown="onGridKey"
        @scroll="scrollTop = ($event.target as HTMLElement).scrollTop"
      >
        <div class="xs-sym-rows" :style="{ height: `${rowCount * cellPx}px` }">
          <div v-for="r in rows" :key="r.row" class="xs-sym-row" :style="{ top: `${r.row * cellPx}px`, height: `${cellPx}px`, gridTemplateColumns: `repeat(${cols}, 1fr)` }">
            <div
              v-for="c in r.codes"
              :id="`${id}-c${c}`"
              :key="c"
              role="option"
              :aria-selected="c === code ? 'true' : 'false'"
              :aria-label="`U+${codeHex(c)}`"
              :class="['xs-sym-cell', c === code ? 'xs-sym-on' : '']"
              :data-xs-sym-code="codeHex(c)"
              @mousedown.prevent
              @click="select(c)"
              @dblclick="select(c), insert()"
            >
              {{ String.fromCodePoint(c) }}
            </div>
          </div>
        </div>
      </div>
      <p>{{ m.symbolRecent }}</p>
      <div class="xs-sym-recent" :style="{ fontFamily: font }">
        <button
          v-for="c in recent"
          :key="c"
          type="button"
          tabindex="-1"
          :aria-label="`U+${codeHex(c)}`"
          :aria-pressed="c === code ? 'true' : 'false'"
          class="xs-sym-cell"
          :data-xs-sym-recent="codeHex(c)"
          @mousedown.prevent
          @click="select(c)"
          @dblclick="select(c), insert()"
        >
          {{ String.fromCodePoint(c) }}
        </button>
      </div>
      <label class="xs-sym-code">
        <span>{{ m.symbolCode }}</span>
        <input class="xs-field xs-field-s" type="text" :value="codeText" maxlength="8" autocomplete="off" autocapitalize="off" spellcheck="false" enterkeyhint="done" data-xs-sym-code-box @input="onCodeInput" />
        <span class="xs-sym-from">{{ m.symbolCodeFrom }}</span>
      </label>
    </div>

    <div v-show="tab === 'special'" class="xs-sym-panel" role="tabpanel" :aria-label="m.symbolTab('special')">
      <p class="xs-sym-head">{{ m.symbolCharacter }}</p>
      <div
        ref="specials"
        class="xs-sym-special"
        role="listbox"
        tabindex="0"
        :aria-label="m.symbolTab('special')"
        :aria-activedescendant="`${id}-s${special}`"
        data-xs-sym-specials
        @keydown="onSpecialKey"
      >
        <div
          v-for="(s, i) in SYMBOL_SPECIALS"
          :id="`${id}-s${i}`"
          :key="s.key"
          role="option"
          :aria-selected="i === special ? 'true' : 'false'"
          :class="['xs-sym-item', i === special ? 'xs-sym-on' : '']"
          :data-xs-sym-special="s.key"
          @mousedown.prevent
          @click="special = i"
          @dblclick="(special = i), insert()"
        >
          <span class="xs-sym-char" :style="{ fontFamily: font }">{{ String.fromCodePoint(s.code) }}</span>
          <span>{{ m.symbolSpecial(s.key) }}</span>
        </div>
      </div>
    </div>

    <div class="xs-dialog-foot">
      <button type="button" class="xs-button xs-button-main" data-xs-sym-insert @click="insert()">{{ m.symbolInsert }}</button>
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="ctx.closeSymbol()">{{ inserted ? m.close : m.cancel }}</button>
    </div>
  </DialogShell>
</template>

<style src="../dialogs.css"></style>
