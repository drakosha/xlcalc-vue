<script setup lang="ts">
// Excel's Conditional Formatting Rules Manager: the rules of the selection or of the sheet in priority order, each with its
// description, format preview, Applies to (editable) and Stop If True; New Rule, Edit Rule, Delete Rule, Duplicate Rule,
// Move Up and Move Down. OK and Apply write the list in one call to the source's `setRules`; Close discards what was not applied.
import { computed, ref, watch } from 'vue'
import type { ConditionalFormat } from 'xlcalc/contract'
import { describeRule } from '../core/cf-describe'
import { appliesText, parseApplies, rectsOf, ruleMeets } from '../core/cf-ranges'
import { useSheetContext } from './context'
import CfPreview from './CfPreview.vue'
import DialogShell from './DialogShell.vue'
import { fieldValue } from './dom'
import '../controls.css'

const ctx = useSheetContext()
const cf = ctx.cf
const m = ctx.messages
const state = computed(() => cf.manager.value)

/** A rule in the list: an id that survives moves, the rule, and its Applies to box while it holds text that is not a range. */
interface Row {
  id: number
  rule: ConditionalFormat
  applies?: string
}
let ids = 0
const rows = ref<Row[]>([])
const original = ref('')
const picked = ref<number | null>(null)
const showFor = ref<'selection' | 'sheet'>('selection')

watch(
  () => state.value?.rules,
  (rules) => {
    if (!rules) return
    rows.value = rules.map((rule) => ({ id: ++ids, rule }))
    original.value = JSON.stringify(rules)
    if (!rows.value.some((r) => r.id === picked.value)) picked.value = null
  },
  { immediate: true },
)

const selection = computed(() => (state.value ? (rectsOf(state.value.ranges)[0] ?? null) : null))
const shown = computed(() => rows.value.filter((r) => showFor.value === 'sheet' || !selection.value || ruleMeets(r.rule, selection.value)))
const current = computed(() => rows.value.find((r) => r.id === picked.value) ?? null)
const separator = computed(() => cf.locale()?.listSeparator ?? ',')
const bad = computed(() => rows.value.some((r) => r.applies !== undefined))
const changed = computed(() => JSON.stringify(rows.value.map((r) => r.rule)) !== original.value)
const describe = (rule: ConditionalFormat) => (state.value ? describeRule(rule, m.value, cf.shown(state.value, rule.ranges)) : '')
const statistic = (rule: ConditionalFormat) => rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet'

/** A typed Applies to: the rule takes the ranges once they parse; until then the box keeps the text and OK waits. */
function setApplies(row: Row, text: string): void {
  const ranges = state.value ? parseApplies(text, separator.value, state.value.sheet) : null
  if (!ranges) {
    row.applies = text
    return
  }
  row.rule = { ...row.rule, ranges }
  delete row.applies
}
function setStop(row: Row, on: boolean): void {
  const { stopIfTrue: _, ...rest } = row.rule
  row.rule = (on ? { ...rest, stopIfTrue: true } : rest) as ConditionalFormat
}

function newRule(): void {
  const at = state.value
  if (!at) return
  cf.openRule({
    ...at,
    base: null,
    done: (rule) => {
      const id = ++ids
      rows.value = [{ id, rule }, ...rows.value]
      picked.value = id
    },
  })
}
function editRule(): void {
  const at = state.value
  const row = current.value
  if (!at || !row) return
  cf.openRule({ ...at, ranges: row.rule.ranges, base: row.rule, done: (rule) => (row.rule = rule) })
}
/** Duplicate Rule: a copy of the rule right after it, picked. */
function duplicateRule(): void {
  const row = current.value
  if (!row) return
  const id = ++ids
  const list = [...rows.value]
  list.splice(list.indexOf(row) + 1, 0, { id, rule: JSON.parse(JSON.stringify(row.rule)) as ConditionalFormat })
  rows.value = list
  picked.value = id
}
function deleteRule(): void {
  const row = current.value
  if (!row) return
  const list = shown.value
  const next = list[list.indexOf(row) + 1] ?? list[list.indexOf(row) - 1]
  rows.value = rows.value.filter((r) => r !== row)
  picked.value = next?.id ?? null
}
/** Move Up and Move Down: the rule changes places with the one next to it in the list shown. */
function move(by: -1 | 1): void {
  const row = current.value
  if (!row) return
  const other = shown.value[shown.value.indexOf(row) + by]
  if (!other) return
  const list = [...rows.value]
  const a = list.indexOf(row)
  const b = list.indexOf(other)
  list[a] = other
  list[b] = row
  rows.value = list
}
const canMove = (by: -1 | 1) => !!current.value && !!shown.value[shown.value.indexOf(current.value) + by]

function apply(close: boolean): void {
  if (bad.value) return
  if (!changed.value) return void (close && cf.managerClose())
  void cf.managerWrite(
    rows.value.map((r) => r.rule),
    close,
  )
}
</script>

<template>
  <DialogShell class="xs-cf-manager" :outside="false" full :aria-label="m.cfManager" data-cf-manager @close="cf.managerClose()">
    <header class="xs-cf-head">
      <span class="xs-modal-title">{{ m.cfManager }}</span>
    </header>
    <div class="xs-cf-body">
      <label class="xs-cf-line">
        <span>{{ m.cfShowFor }}</span>
        <select v-model="showFor" class="xs-field" data-cf-show>
          <option value="selection">{{ m.cfShowSelection }}</option>
          <option value="sheet">{{ m.cfShowSheet }}</option>
        </select>
      </label>
      <div class="xs-cf-tools">
        <button type="button" class="xs-button" :disabled="!state?.rules" data-cf-new @click="newRule">{{ m.cfNewRule }}</button>
        <button type="button" class="xs-button" :disabled="!current" data-cf-edit @click="editRule">{{ m.cfEditRule }}</button>
        <button type="button" class="xs-button" :disabled="!current" data-cf-delete @click="deleteRule">{{ m.cfDeleteRule }}</button>
        <button type="button" class="xs-button" :disabled="!current" data-cf-duplicate @click="duplicateRule">{{ m.cfDuplicateRule }}</button>
        <button type="button" class="xs-button" :disabled="!canMove(-1)" :title="m.cfMoveUp" data-cf-up @click="move(-1)">▲ {{ m.cfMoveUp }}</button>
        <button type="button" class="xs-button" :disabled="!canMove(1)" :title="m.cfMoveDown" data-cf-down @click="move(1)">▼ {{ m.cfMoveDown }}</button>
      </div>
      <div class="xs-cf-table" role="grid" :aria-label="m.cfManager">
        <div class="xs-cf-tr xs-cf-th" role="row">
          <span role="columnheader">{{ m.cfColRule }}</span>
          <span role="columnheader">{{ m.cfColFormat }}</span>
          <span role="columnheader">{{ m.cfColApplies }}</span>
          <span role="columnheader">{{ m.cfColStop }}</span>
        </div>
        <p v-if="!state?.rules" class="xs-muted xs-cf-empty">{{ m.cfLoading }}</p>
        <div
          v-for="r in shown"
          :key="r.id"
          class="xs-cf-tr"
          role="row"
          :aria-selected="r.id === picked"
          :data-cf-row="rows.indexOf(r)"
          @click="picked = r.id"
          @dblclick="(picked = r.id), editRule()"
        >
          <span role="gridcell" class="xs-cf-desc-text" data-cf-desc>{{ describe(r.rule) }}</span>
          <span role="gridcell"><CfPreview :rule="r.rule" :messages="m" /></span>
          <span role="gridcell">
            <input
              class="xs-field xs-cf-applies"
              type="text"
              spellcheck="false"
              autocomplete="off"
              :value="r.applies ?? appliesText(r.rule.ranges, separator)"
              :aria-invalid="r.applies !== undefined"
              :aria-label="m.cfColApplies"
              data-cf-applies
              @focus="picked = r.id"
              @change="setApplies(r, fieldValue($event))"
            />
          </span>
          <span role="gridcell">
            <input type="checkbox" :checked="!!r.rule.stopIfTrue" :disabled="statistic(r.rule)" :aria-label="m.cfColStop" data-cf-stop @change="setStop(r, ($event.target as HTMLInputElement).checked)" />
          </span>
        </div>
      </div>
      <p v-if="bad" class="xs-cf-error" role="alert">{{ m.cfAppliesBad }}</p>
    </div>
    <footer class="xs-cf-foot">
      <button type="button" class="xs-button xs-cf-btn" data-cf-close @click="cf.managerClose()">{{ m.close }}</button>
      <button type="button" class="xs-button xs-cf-btn" :disabled="bad || !changed" data-cf-apply @click="apply(false)">{{ m.cfApply }}</button>
      <button type="button" class="xs-button xs-button-main xs-cf-btn" :disabled="bad || !state?.rules" data-cf-ok @click="apply(true)">{{ m.ok }}</button>
    </footer>
  </DialogShell>
</template>

<style src="../cf.css"></style>
