<script setup lang="ts">
// Excel's Data Validation dialog for the selection: Settings (Allow, Data and its boxes by type and comparison, Ignore blank,
// In-cell dropdown, Apply these changes to all other cells with the same settings), Input Message and Error Alert, Clear All.
// It opens with the rule the sheet found for the active cell, its formulas with `=`; OK reads the boxes as Excel does, shows
// Excel's message for a box it refuses, and writes the rule through the sheet; Cancel closes. Full screen on a phone.
import { computed, ref, toRaw, watch } from 'vue'
import { boxesOf, emptyDraft, isCompared, ruleOf, VALIDATION_LIMITS, VALIDATION_OPERATORS, VALIDATION_TABS, VALIDATION_TYPES, type DraftError, type ValidationDraft, type ValidationTab } from '../validation-rules'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { IconClose } from './icons'
import '../controls.css'
import '../format-dialog.css'

const ctx = useSheetContext()
const m = ctx.messages
const dialog = ctx.rules

const tab = ref<ValidationTab>('settings')
const draft = ref<ValidationDraft | null>(null)
const error = ref<string | null>(null)
const busy = ref(false)

watch(
  () => dialog.state.value,
  (d) => {
    tab.value = 'settings'
    error.value = null
    busy.value = false
    draft.value = d ? structuredClone(d.draft) : null
  },
  { immediate: true },
)

const boxes = computed(() => (draft.value ? boxesOf(draft.value.type, draft.value.operator) : []))
/** The option applies only when the dialog opened with a rule's settings, as in Excel. */
const canApplyAll = computed(() => !!dialog.state.value?.shown)

function errorText(e: DraftError): string {
  const name = (b: Parameters<typeof m.value.dvBox>[0]) => m.value.dvBox(b)
  if (e.kind === 'order') return m.value.dvOrder(name(e.min), name(e.max))
  return e.kind === 'empty' ? m.value.dvEmpty(name(e.box)) : m.value.dvInvalid(name(e.box))
}

/** Clear All: Excel's defaults on every tab; the dialog stays open, and OK removes the rule from the cells. */
function clearAll(): void {
  const keep = draft.value?.applyAll ?? false
  draft.value = { ...emptyDraft(), applyAll: keep }
  error.value = null
}

async function ok(): Promise<void> {
  const d = dialog.state.value
  if (!d || !draft.value || busy.value) return
  // plain objects: a source in a worker thread clones them, and a reactive proxy cannot be cloned
  const made = ruleOf(toRaw(draft.value), d.text)
  if (!made.ok) {
    tab.value = 'settings'
    error.value = errorText(made.error)
    return
  }
  busy.value = true
  error.value = null
  const refused = await dialog.apply(made.rule, canApplyAll.value && draft.value.applyAll)
  busy.value = false
  if (refused) error.value = refused
}
const cancel = () => dialog.close()
</script>

<template>
  <DialogShell
    v-if="draft && dialog.state.value"
    as="form"
    class="xs-fmt xs-dv"
    full
    focus="[data-dv-type]"
    :aria-label="m.dataValidation"
    data-dv
    novalidate
    @submit.prevent="ok"
    @close="cancel"
  >
    <header class="xs-fmt-head">
      <span class="xs-modal-title xs-fmt-title">{{ m.dataValidation }}</span>
      <button type="button" class="xs-tool xs-fmt-x" :title="m.cancel" :aria-label="m.cancel" @click="cancel">
        <IconClose />
      </button>
    </header>
    <div class="xs-fmt-tabs" role="tablist">
      <button v-for="t in VALIDATION_TABS" :key="t" type="button" class="xs-fmt-tab" role="tab" :aria-selected="tab === t" :data-dv-tab="t" @click="tab = t">
        {{ m.dvTab(t) }}
      </button>
    </div>

    <div class="xs-fmt-body" role="tabpanel" :aria-label="m.dvTab(tab)">
      <div v-show="tab === 'settings'" class="xs-fmt-stack">
        <fieldset class="xs-fmt-box">
          <legend>{{ m.dvCriteria }}</legend>
          <div class="xs-fmt-grid">
            <label class="xs-fmt-field">
              <span>{{ m.dvAllow }}</span>
              <select v-model="draft.type" class="xs-field xs-fmt-select" data-dv-type>
                <option v-for="t in VALIDATION_TYPES" :key="t" :value="t">{{ m.dvType(t) }}</option>
              </select>
            </label>
            <label class="xs-fmt-check xs-dv-check">
              <input v-model="draft.allowBlank" type="checkbox" :disabled="draft.type === 'any'" data-dv-blank />
              {{ m.dvIgnoreBlank }}
            </label>
            <label v-if="isCompared(draft.type)" class="xs-fmt-field">
              <span>{{ m.dvData }}</span>
              <select v-model="draft.operator" class="xs-field xs-fmt-select" data-dv-operator>
                <option v-for="op in VALIDATION_OPERATORS" :key="op" :value="op">{{ m.dvOperator(op) }}</option>
              </select>
            </label>
            <label v-if="draft.type === 'list'" class="xs-fmt-check xs-dv-check">
              <input v-model="draft.dropdown" type="checkbox" data-dv-dropdown />
              {{ m.dvDropdown }}
            </label>
            <label v-for="(box, i) in boxes" :key="box" class="xs-fmt-field xs-fmt-wide">
              <span>{{ m.dvBox(box) }}</span>
              <input
                v-if="i === 0"
                v-model="draft.first"
                class="xs-field xs-fmt-input"
                type="text"
                spellcheck="false"
                autocomplete="off"
                :data-dv-box="box"
              />
              <input v-else v-model="draft.second" class="xs-field xs-fmt-input" type="text" spellcheck="false" autocomplete="off" :data-dv-box="box" />
            </label>
          </div>
        </fieldset>
        <label class="xs-fmt-check">
          <input v-model="draft.applyAll" type="checkbox" :disabled="!canApplyAll" data-dv-all />
          {{ m.dvApplyAll }}
        </label>
      </div>

      <div v-show="tab === 'input'" class="xs-fmt-stack">
        <label class="xs-fmt-check">
          <input v-model="draft.showPrompt" type="checkbox" data-dv-show-prompt />
          {{ m.dvShowPrompt }}
        </label>
        <p class="xs-dv-lead">{{ m.dvPromptLead }}</p>
        <label class="xs-fmt-field">
          <span>{{ m.dvFieldTitle }}</span>
          <input v-model="draft.promptTitle" class="xs-field xs-fmt-input" type="text" :maxlength="VALIDATION_LIMITS.promptTitle" data-dv-prompt-title />
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.dvPromptText }}</span>
          <textarea v-model="draft.prompt" class="xs-field xs-dv-text" rows="5" :maxlength="VALIDATION_LIMITS.prompt" data-dv-prompt />
        </label>
      </div>

      <div v-show="tab === 'error'" class="xs-fmt-stack">
        <label class="xs-fmt-check">
          <input v-model="draft.showError" type="checkbox" data-dv-show-error />
          {{ m.dvShowError }}
        </label>
        <p class="xs-dv-lead">{{ m.dvErrorLead }}</p>
        <div class="xs-fmt-grid">
          <label class="xs-fmt-field">
            <span>{{ m.dvStyle }}</span>
            <select v-model="draft.errorStyle" class="xs-field xs-fmt-select" data-dv-style>
              <option v-for="st in ['stop', 'warning', 'information'] as const" :key="st" :value="st">{{ m.validationStyle(st) }}</option>
            </select>
          </label>
          <label class="xs-fmt-field">
            <span>{{ m.dvFieldTitle }}</span>
            <input v-model="draft.errorTitle" class="xs-field xs-fmt-input" type="text" :maxlength="VALIDATION_LIMITS.errorTitle" data-dv-error-title />
          </label>
          <label class="xs-fmt-field xs-fmt-wide">
            <span>{{ m.dvErrorText }}</span>
            <textarea v-model="draft.error" class="xs-field xs-dv-text" rows="5" :maxlength="VALIDATION_LIMITS.error" data-dv-error />
          </label>
        </div>
      </div>
      <p v-if="error" class="xs-fmt-error xs-dv-error" role="alert" data-dv-message>{{ error }}</p>
    </div>

    <footer class="xs-fmt-foot">
      <button type="button" class="xs-button xs-fmt-btn xs-dv-clear" data-dv-clear @click="clearAll">{{ m.dvClearAll }}</button>
      <button type="button" class="xs-button xs-fmt-btn" data-dv-cancel @click="cancel">{{ m.cancel }}</button>
      <button type="submit" class="xs-button xs-button-main xs-fmt-btn" :disabled="busy" data-dv-ok>{{ m.ok }}</button>
    </footer>
  </DialogShell>
</template>

<style src="../validation-dialog.css"></style>
