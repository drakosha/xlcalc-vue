<script setup lang="ts">
// Excel's Insert Function window: a search, the category box (recently used, all, the source's categories), the functions, and the syntax and description of the one selected.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue'
import { CATEGORY_ORDER } from '../function-categories'
import type { SheetFunctionSignature } from '../source'
import { useSheetContext } from './context'

const emit = defineEmits<{ pick: [name: string]; close: [] }>()
const ctx = useSheetContext()
const m = ctx.messages

/** The functions used last, newest first, as Excel's Most Recently Used; in the sheet's memory. */
const RECENT = 'recentFunctions'
const RECENT_MAX = 10

const names = shallowRef<string[]>([])
const sigs = shallowRef(new Map<string, SheetFunctionSignature | null>())
const recent = (ctx.memory.recall<string[]>(RECENT) ?? []).filter((n) => typeof n === 'string')
const category = ref<string>(recent.length ? 'recent' : 'all')
const query = ref('')
const index = ref(0)
const search = ref<HTMLInputElement | null>(null)
const list = ref<HTMLElement | null>(null)
const ctrl = new AbortController()

/** The categories the source's functions fall in, in Excel's order, by their names in the locale. */
const categories = computed(() => {
  const named = new Map<string, string>()
  for (const sig of sigs.value.values()) if (sig && !named.has(sig.category)) named.set(sig.category, sig.categoryName)
  return CATEGORY_ORDER.filter((c) => named.has(c)).map((c) => ({ key: c, name: named.get(c)! }))
})
/** The category's functions whose names hold the search, in any case: the name itself first, then those it begins, then the rest. */
const shown = computed(() => {
  const known = new Set(names.value)
  const base = category.value === 'recent' ? recent.filter((n) => known.has(n)) : category.value === 'all' ? names.value : names.value.filter((n) => sigs.value.get(n)?.category === category.value)
  const q = query.value.trim().toUpperCase()
  if (!q) return base
  const hits: { name: string; rank: number }[] = []
  for (const name of base) {
    const upper = name.toUpperCase()
    if (upper.includes(q)) hits.push({ name, rank: upper === q ? 0 : upper.startsWith(q) ? 1 : 2 })
  }
  return hits.sort((a, b) => a.rank - b.rank).map((x) => x.name)
})
const selected = computed(() => shown.value[index.value] ?? null)
const selectedSig = computed(() => (selected.value ? (sigs.value.get(selected.value) ?? null) : null))

watch([category, query], () => (index.value = 0))
watch(query, (q) => {
  if (q.trim() && category.value === 'recent') category.value = 'all'
})
watch(index, async () => {
  await nextTick()
  list.value?.querySelector('.xs-ifn-on')?.scrollIntoView({ block: 'nearest' })
})

onMounted(async () => {
  search.value?.focus()
  const src = ctx.source()
  if (!src.functionNames) return
  const got = await src.functionNames(ctrl.signal).catch(() => null)
  if (!got || ctrl.signal.aborted) return
  names.value = got
  // Every name's signature, once per source (the tip keeps them): the categories, the syntax and the description.
  const all = await Promise.all(got.map(async (name) => [name, await ctx.ed.tip.signature(name)] as const))
  if (!ctrl.signal.aborted) sigs.value = new Map(all)
})
onBeforeUnmount(() => ctrl.abort())

function ok(): void {
  const name = selected.value
  if (!name) return
  ctx.memory.remember(RECENT, [name, ...recent.filter((n) => n !== name)].slice(0, RECENT_MAX))
  emit('pick', name)
}
function onKey(e: KeyboardEvent): void {
  const last = shown.value.length - 1
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault()
    index.value = Math.max(0, Math.min(last, index.value + (e.key === 'ArrowDown' ? 1 : -1)))
  } else if (e.key === 'Enter') {
    e.preventDefault()
    ok()
  } else if (e.key === 'Escape') {
    e.preventDefault()
    e.stopPropagation()
    emit('close')
  }
}
</script>

<template>
  <div class="xs-ifn-back" data-xs-keep-edit @pointerdown.self="emit('close')">
    <div class="xs-ifn" role="dialog" aria-modal="true" :aria-label="m.insertFunctionTitle" tabindex="-1" data-xs-insert-function @keydown="onKey">
      <p class="xs-ifn-title">{{ m.insertFunctionTitle }}</p>
      <input
        ref="search"
        v-model="query"
        class="xs-ifn-input"
        type="text"
        enterkeyhint="search"
        autocomplete="off"
        spellcheck="false"
        :placeholder="m.insertFunctionSearch"
        :aria-label="m.insertFunctionSearch"
        data-xs-if-search
      />
      <label class="xs-ifn-field">
        <span>{{ m.insertFunctionCategory }}</span>
        <select v-model="category" class="xs-ifn-input" data-xs-if-category>
          <option value="recent">{{ m.insertFunctionRecent }}</option>
          <option value="all">{{ m.insertFunctionAll }}</option>
          <option v-for="c in categories" :key="c.key" :value="c.key">{{ c.name }}</option>
        </select>
      </label>
      <p class="xs-ifn-label">{{ m.insertFunctionPick }}</p>
      <div ref="list" class="xs-ifn-list" role="listbox" :aria-label="m.insertFunctionPick">
        <div
          v-for="(name, i) in shown"
          :key="name"
          role="option"
          :aria-selected="i === index ? 'true' : 'false'"
          :class="['xs-ifn-item', i === index ? 'xs-ifn-on' : '']"
          data-xs-if-item
          @mousedown.prevent
          @click="index = i"
          @dblclick="(index = i), ok()"
        >
          {{ name }}
        </div>
      </div>
      <p class="xs-ifn-syntax" data-xs-if-syntax>{{ selectedSig?.text ?? selected ?? '' }}</p>
      <p class="xs-ifn-about" data-xs-if-about>{{ selectedSig?.description ?? '' }}</p>
      <div class="xs-dialog-foot">
        <button type="button" class="xs-button" data-xs-dialog-cancel @click="emit('close')">{{ m.cancel }}</button>
        <button type="button" class="xs-button xs-button-main" :disabled="!selected" data-xs-dialog-ok @click="ok()">{{ m.ok }}</button>
      </div>
    </div>
  </div>
</template>

<style src="../insert-function.css"></style>
