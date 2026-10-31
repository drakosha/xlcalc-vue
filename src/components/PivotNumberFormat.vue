<script setup lang="ts">
// Number Format… of a value field: Format Cells with its Number tab alone, Excel's Number Format list in the workbook's
// locale and the format code; OK gives the value field its format (from Value Field Settings, the dialog's draft).
import { computed, ref, watch } from 'vue'
import type { NumberFormat } from 'xlcalc/contract'
import { useNumberFormats } from '../number-formats'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { IconClose } from './icons'
import '../format-dialog.css'
import '../pivot-dialogs.css'

const ctx = useSheetContext()
const m = ctx.messages
const dialogs = ctx.pivotDialogs
const numbers = useNumberFormats(() => ctx.locale.value)
const d = dialogs.state
const shown = computed(() => (d.value?.kind === 'value' && d.value.format) || d.value?.kind === 'numberFormat')
/** The code in the box, the format it stands for (a built-in one by its id, as the list gives it), and the code it opened with. */
const code = ref('General')
const format = ref<NumberFormat>('General')
const start = ref('General')
/** A format as the box shows it: a built-in id by its code, as the source gives it; General without one. */
function codeOf(fmt: NumberFormat | undefined, index: number): string {
  if (fmt === undefined) return 'General'
  if (typeof fmt === 'string') return fmt
  return numbers.choices.value.find((c) => c.format === fmt)?.code ?? d.value?.t.cells.formats[index] ?? String(fmt)
}
watch(
  [shown, () => numbers.choices.value.length],
  ([on]) => {
    const x = d.value
    if (!on || !x) return
    const fmt = x.kind === 'value' ? x.draft.numberFormat : x.kind === 'numberFormat' ? x.format : undefined
    format.value = fmt ?? 'General'
    code.value = start.value = codeOf(fmt, 'index' in x ? x.index : 0)
  },
  { immediate: true },
)
const category = computed(() => numbers.categoryOf(code.value.trim() || undefined))
function pick(c: { code: string; format: NumberFormat }): void {
  code.value = c.code
  format.value = c.format
}
function type(text: string): void {
  code.value = text
  format.value = text.trim()
}
function cancel(): void {
  const x = d.value
  if (x?.kind === 'value') x.format = false
  else dialogs.close()
}
const ok = () => code.value.trim() && void dialogs.applyNumberFormat(code.value.trim() === start.value ? null : format.value)
</script>

<template>
  <DialogShell v-if="shown" as="form" class="xs-fmt xs-pnf" full focus="[data-pnf-code]" :aria-label="m.formatCells" data-pivot-number-format novalidate @submit.prevent="ok" @close="cancel">
    <header class="xs-fmt-head">
      <span class="xs-modal-title xs-fmt-title">{{ m.formatCells }}</span>
      <button type="button" class="xs-tool xs-fmt-x" :title="m.cancel" :aria-label="m.cancel" @click="cancel"><IconClose /></button>
    </header>
    <div class="xs-fmt-tabs" role="tablist">
      <button type="button" class="xs-fmt-tab" role="tab" aria-selected="true">{{ m.fmtTab('number') }}</button>
    </div>
    <div class="xs-fmt-body xs-fmt-stack" role="tabpanel">
      <p class="xs-fmt-category">{{ m.fmtCategory }}: <b>{{ category?.name ?? '' }}</b></p>
      <fieldset class="xs-fmt-box">
        <legend>{{ m.numberFormat }}</legend>
        <div class="xs-fmt-list" role="radiogroup" :aria-label="m.numberFormat">
          <label v-for="c in numbers.choices.value" :key="c.key" class="xs-fmt-choice">
            <input type="radio" name="xs-pnf-number" :checked="code === c.code" :data-pnf-number="c.key" @change="pick(c)" />
            {{ c.name }}
          </label>
        </div>
      </fieldset>
      <label class="xs-fmt-field">
        <span>{{ m.fmtCode }}</span>
        <input :value="code" class="xs-field xs-fmt-input" type="text" spellcheck="false" autocomplete="off" :aria-invalid="!code.trim()" data-pnf-code @input="type(($event.target as HTMLInputElement).value)" />
      </label>
    </div>
    <footer class="xs-fmt-foot">
      <button type="submit" class="xs-button xs-button-main xs-fmt-btn" :disabled="!code.trim()" data-xs-dialog-ok>{{ m.ok }}</button>
      <button type="button" class="xs-button xs-fmt-btn" data-xs-dialog-cancel @click="cancel">{{ m.cancel }}</button>
    </footer>
  </DialogShell>
</template>
