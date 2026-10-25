<script setup lang="ts">
// A quick rule's dialog from Excel's Highlight Cells Rules and Top/Bottom Rules: its value boxes (a value, a reference or a
// formula after `=`, typed for the active cell), a text, a period, Duplicate or Unique, or a rank, and the format "with"
// which the cells are shown: Excel's presets, or Custom Format in the rule's Format Cells.
import { computed, ref, shallowRef } from 'vue'
import type { ConditionalStyle } from 'xlcalc/contract'
import { boxFormula, textRuleFormula } from '../core/cf-formula'
import { PRESET_STYLES, rankMax, STYLE_PRESETS, type QuickInput, type StylePreset } from '../core/cf-presets'
import { PERIODS, type Period } from '../core/cf-draft'
import { loadEntry } from '../entry'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { fieldValue } from './dom'
import '../controls.css'

const ctx = useSheetContext()
const cf = ctx.cf
const m = ctx.messages
const target = cf.quick.value!
const kind = target.kind

const boxes = kind === 'between' ? 2 : kind === 'greaterThan' || kind === 'lessThan' || kind === 'equalTo' ? 1 : 0
const max = rankMax(kind)
const values = ref(['', ''])
const text = ref('')
const period = ref<Period>('yesterday')
const unique = ref(false)
const rank = ref('10')
const preset = ref<StylePreset | 'custom'>('lightRedDarkRed')
const custom = shallowRef<ConditionalStyle | null>(null)
const error = ref('')
let lastPreset: StylePreset = 'lightRedDarkRed'

/** The words after the inputs: "with", "values with", "% with", "for the selected range with". */
const lead = computed(() => (kind === 'duplicateValues' ? m.value.cfValuesWith : kind === 'aboveAverage' || kind === 'belowAverage' ? m.value.cfForRange : m.value.cfWith))

/** Custom Format opens the rule's Format Cells; Cancel there goes back to the preset chosen before. */
function pickStyle(v: string): void {
  if (v !== 'custom') {
    preset.value = v as StylePreset
    lastPreset = preset.value
    return
  }
  preset.value = 'custom'
  cf.openStyle({
    style: custom.value ?? {},
    done: (st) => (custom.value = st),
    cancel: () => {
      if (!custom.value) preset.value = lastPreset
    },
  })
}

async function ok(): Promise<void> {
  error.value = ''
  const entry = await loadEntry().catch(() => null)
  if (!entry) return
  const locale = cf.locale()
  const q: QuickInput = { style: preset.value === 'custom' ? (custom.value ?? {}) : PRESET_STYLES[preset.value] }
  if (boxes) {
    const formulas: string[] = []
    for (const v of values.value.slice(0, boxes)) {
      const got = boxFormula(v, entry, locale)
      if (!got.ok) return void (error.value = got.message ?? m.value.cfValueMissing)
      formulas.push(got.formula)
    }
    q.formulas = formulas
  }
  if (kind === 'textContains') {
    const t = text.value.trim()
    if (!t) return void (error.value = m.value.cfValueMissing)
    if (t.startsWith('=')) {
      const got = boxFormula(t, entry, locale)
      if (!got.ok) return void (error.value = got.message ?? m.value.cfValueMissing)
      q.text = got.formula
      q.textFormula = textRuleFormula('containsText', got.formula, cf.activeAddress(target), locale)
    } else q.text = text.value
  }
  if (kind === 'dateOccurring') q.period = period.value
  if (kind === 'duplicateValues') q.unique = unique.value
  if (max !== null) {
    const n = /^\d+$/.test(rank.value.trim()) ? Number(rank.value.trim()) : NaN
    if (!(n >= 1 && n <= max)) return void (error.value = m.value.cfRankRange(max))
    q.rank = n
  }
  cf.quickOk(q)
}
</script>

<template>
  <DialogShell as="form" class="xs-cf-quick" focus="input, select" :aria-label="m.cfQuickTitle(kind)" :data-cf-quick="kind" novalidate @submit.prevent="ok" @close="cf.quickClose()">
    <header class="xs-cf-head">
      <span class="xs-modal-title">{{ m.cfQuickTitle(kind) }}</span>
    </header>
    <div class="xs-cf-body">
      <p>{{ m.cfQuickLabel(kind) }}</p>
      <div class="xs-cf-row">
        <template v-if="boxes">
          <input v-model="values[0]" class="xs-field xs-cf-box" type="text" spellcheck="false" autocomplete="off" data-cf-value="0" />
          <template v-if="boxes === 2">
            <span>{{ m.cfAnd }}</span>
            <input v-model="values[1]" class="xs-field xs-cf-box" type="text" spellcheck="false" autocomplete="off" data-cf-value="1" />
          </template>
        </template>
        <input v-else-if="kind === 'textContains'" v-model="text" class="xs-field xs-cf-box" type="text" spellcheck="false" autocomplete="off" data-cf-text />
        <select v-else-if="kind === 'dateOccurring'" v-model="period" class="xs-field xs-cf-box" data-cf-period>
          <option v-for="p in PERIODS" :key="p" :value="p">{{ m.cfPeriod(p) }}</option>
        </select>
        <select v-else-if="kind === 'duplicateValues'" v-model="unique" class="xs-field" data-cf-unique>
          <option :value="false">{{ m.cfDuplicate(false) }}</option>
          <option :value="true">{{ m.cfDuplicate(true) }}</option>
        </select>
        <template v-else-if="max !== null">
          <input v-model="rank" class="xs-field xs-cf-rank" type="text" inputmode="numeric" data-cf-rank />
          <span v-if="max === 100">%</span>
        </template>
        <span>{{ lead }}</span>
        <select class="xs-field xs-cf-with" :value="preset" data-cf-preset @change="pickStyle(fieldValue($event))">
          <option v-for="p in STYLE_PRESETS" :key="p" :value="p">{{ m.cfPreset(p) }}</option>
          <option value="custom">{{ m.cfCustomFormat }}</option>
        </select>
      </div>
      <p v-if="error" class="xs-cf-error" role="alert">{{ error }}</p>
    </div>
    <footer class="xs-cf-foot">
      <button type="button" class="xs-button xs-cf-btn" data-cf-cancel @click="cf.quickClose()">{{ m.cancel }}</button>
      <button type="submit" class="xs-button xs-button-main xs-cf-btn" data-cf-ok>{{ m.ok }}</button>
    </footer>
  </DialogShell>
</template>

<style src="../cf.css"></style>
