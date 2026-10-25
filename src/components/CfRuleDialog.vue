<script setup lang="ts">
// Excel's New Formatting Rule and Edit Formatting Rule: a rule type from Excel's list and its description, the preview and
// Format… (the rule's Format Cells). Formulas are shown and typed for the active cell, with `=`, in the workbook's locale;
// OK writes them for the rule's first cell and passes the rule on.
import { computed, ref } from 'vue'
import { AVERAGE_KINDS, CONTAIN_KINDS, draftOf, newDraft, OPERATORS, PERIODS, previewOf, RULE_KINDS, ruleOf, TEXT_OPS, type DraftError } from '../core/cf-draft'
import { loadEntry } from '../entry'
import { useSheetContext } from './context'
import CfPreview from './CfPreview.vue'
import CfValuesFields from './CfValuesFields.vue'
import DialogShell from './DialogShell.vue'
import '../controls.css'

const ctx = useSheetContext()
const cf = ctx.cf
const m = ctx.messages
const dialog = cf.rule.value!
const draft = ref(dialog.base ? draftOf(dialog.base, cf.shown(dialog, dialog.ranges)) : newDraft())
const error = ref('')
const preview = computed(() => previewOf(draft.value))
const twoValues = computed(() => draft.value.operator === 'between' || draft.value.operator === 'notBetween')

function errorText(e: DraftError): string {
  const t = m.value
  if (e.key === 'relative') return t.cfRelative
  if (e.key === 'percent') return t.cfPercentRange
  if (e.key === 'rank') return t.cfRankRange(e.max)
  if (e.key === 'formula' && e.message) return e.message
  return t.cfValueMissing
}

function format(): void {
  cf.openStyle({ style: draft.value.style, done: (st) => (draft.value.style = st) })
}

async function ok(): Promise<void> {
  error.value = ''
  const entry = await loadEntry().catch(() => null)
  if (!entry) return
  const got = ruleOf(draft.value, dialog.base, { ranges: dialog.ranges, locale: cf.locale(), entry, stored: cf.stored(dialog), active: cf.activeAddress(dialog) })
  if (!got.ok) return void (error.value = errorText(got.error))
  cf.ruleClose()
  dialog.done(got.rule)
}
</script>

<template>
  <DialogShell
    as="form"
    class="xs-cf-rule"
    :outside="false"
    full
    focus="[aria-selected='true']"
    :aria-label="dialog.base ? m.cfEditTitle : m.cfNewTitle"
    data-cf-rule
    novalidate
    @submit.prevent="ok"
    @close="cf.ruleClose()"
  >
    <header class="xs-cf-head">
      <span class="xs-modal-title">{{ dialog.base ? m.cfEditTitle : m.cfNewTitle }}</span>
    </header>
    <div class="xs-cf-body">
      <p>{{ m.cfSelectType }}</p>
      <div class="xs-cf-types" role="listbox" :aria-label="m.cfSelectType">
        <button
          v-for="k in RULE_KINDS"
          :key="k"
          type="button"
          class="xs-cf-type"
          role="option"
          :aria-selected="draft.kind === k"
          :data-cf-kind="k"
          @click="draft.kind = k"
        >
          {{ m.cfType(k) }}
        </button>
      </div>
      <p>{{ m.cfEditDescription }}</p>
      <div class="xs-cf-desc">
        <CfValuesFields v-if="draft.kind === 'values'" :draft="draft" :messages="m" />

        <template v-else-if="draft.kind === 'contain'">
          <p>{{ m.cfFormatOnly }}</p>
          <div class="xs-cf-row">
            <select v-model="draft.contain" class="xs-field" data-cf-contain>
              <option v-for="k in CONTAIN_KINDS" :key="k" :value="k">{{ m.cfContain(k) }}</option>
            </select>
            <template v-if="draft.contain === 'cellValue'">
              <select v-model="draft.operator" class="xs-field" data-cf-operator>
                <option v-for="o in OPERATORS" :key="o" :value="o">{{ m.cfOperator(o) }}</option>
              </select>
              <input v-model="draft.value1" class="xs-field xs-cf-box" type="text" spellcheck="false" autocomplete="off" data-cf-value="0" />
              <template v-if="twoValues">
                <span>{{ m.cfAnd }}</span>
                <input v-model="draft.value2" class="xs-field xs-cf-box" type="text" spellcheck="false" autocomplete="off" data-cf-value="1" />
              </template>
            </template>
            <template v-else-if="draft.contain === 'text'">
              <select v-model="draft.textOp" class="xs-field" data-cf-text-op>
                <option v-for="o in TEXT_OPS" :key="o" :value="o">{{ m.cfTextOp(o) }}</option>
              </select>
              <input v-model="draft.text" class="xs-field xs-cf-box" type="text" spellcheck="false" autocomplete="off" data-cf-text />
            </template>
            <select v-else-if="draft.contain === 'dates'" v-model="draft.period" class="xs-field" data-cf-period>
              <option v-for="p in PERIODS" :key="p" :value="p">{{ m.cfPeriod(p) }}</option>
            </select>
          </div>
        </template>

        <template v-else-if="draft.kind === 'rank'">
          <p>{{ m.cfRankLabel }}</p>
          <div class="xs-cf-row">
            <select v-model="draft.top" class="xs-field" data-cf-top>
              <option :value="true">{{ m.cfTopBottom(true) }}</option>
              <option :value="false">{{ m.cfTopBottom(false) }}</option>
            </select>
            <input v-model="draft.rank" class="xs-field xs-cf-rank" type="text" inputmode="numeric" data-cf-rank />
            <label class="xs-fmt-check">
              <input v-model="draft.percent" type="checkbox" data-cf-percent />
              {{ m.cfPercentOf }}
            </label>
          </div>
        </template>

        <template v-else-if="draft.kind === 'average'">
          <p>{{ m.cfAverageLabel }}</p>
          <div class="xs-cf-row">
            <select v-model="draft.average" class="xs-field" data-cf-average>
              <option v-for="k in AVERAGE_KINDS" :key="k" :value="k">{{ m.cfAverageKind(k) }}</option>
            </select>
            <span>{{ m.cfAverageOf }}</span>
          </div>
        </template>

        <template v-else-if="draft.kind === 'unique'">
          <p>{{ m.cfUniqueLabel }}</p>
          <div class="xs-cf-row">
            <select v-model="draft.unique" class="xs-field" data-cf-unique>
              <option :value="false">{{ m.cfDuplicate(false) }}</option>
              <option :value="true">{{ m.cfDuplicate(true) }}</option>
            </select>
            <span>{{ m.cfInRange }}</span>
          </div>
        </template>

        <template v-else>
          <p>{{ m.cfFormulaLabel }}</p>
          <input v-model="draft.formula" class="xs-field xs-cf-formula" type="text" spellcheck="false" autocomplete="off" data-cf-formula />
        </template>

        <div class="xs-cf-row xs-cf-preview-row">
          <span>{{ m.cfPreview }}</span>
          <CfPreview :rule="preview" :messages="m" />
          <button v-if="draft.kind !== 'values'" type="button" class="xs-button" data-cf-format @click="format">{{ m.cfFormatButton }}</button>
        </div>
      </div>
      <p v-if="error" class="xs-cf-error" role="alert">{{ error }}</p>
    </div>
    <footer class="xs-cf-foot">
      <button type="button" class="xs-button xs-cf-btn" data-cf-cancel @click="cf.ruleClose()">{{ m.cancel }}</button>
      <button type="submit" class="xs-button xs-button-main xs-cf-btn" data-cf-ok>{{ m.ok }}</button>
    </footer>
  </DialogShell>
</template>

<style src="../cf.css"></style>
