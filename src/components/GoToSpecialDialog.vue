<script setup lang="ts">
// Excel's Go To Special dialog in its two columns, with the options the component can find (`useSpecial`): Constants and
// Formulas by the kinds of their values, Blanks, Current region; Last cell, Visible cells only, Conditional formats and Data
// validation, each with All or Same. Excel's Comments, Current array, Objects, Row and Column differences, Precedents and
// Dependents are left out: no source method tells them.
import { computed, ref, useId } from 'vue'
import type { SpecialValue } from 'xlcalc/contract'
import type { SheetMessages } from '../contract/messages'
import { SPECIAL_VALUES, type SpecialKind } from '../core/special'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'

const ctx = useSheetContext()
const m = ctx.messages
const special = ctx.special
const titleId = `${useId()}-title`
const text = (key: keyof SheetMessages) => m.value[key] as string

const LEFT: { kind: SpecialKind; label: keyof SheetMessages }[] = [
  { kind: 'constants', label: 'specialConstants' },
  { kind: 'formulas', label: 'specialFormulas' },
]
const VALUE_LABELS: Record<SpecialValue, keyof SheetMessages> = { numbers: 'specialNumbers', text: 'specialText', logicals: 'specialLogicals', errors: 'specialErrors' }
const RIGHT: { kind: SpecialKind; label: keyof SheetMessages }[] = [
  { kind: 'lastCell', label: 'specialLastCell' },
  { kind: 'visible', label: 'specialVisible' },
  { kind: 'conditionalFormats', label: 'specialConditional' },
  { kind: 'validation', label: 'specialValidation' },
]

const start = special.last.value
const kind = ref<SpecialKind>(start.kind)
const values = ref<SpecialValue[]>([...start.values])
const same = ref(start.same)
/** The kinds of values apply to Constants and Formulas, All and Same to the rules, as Excel enables them. */
const byValue = computed(() => kind.value === 'constants' || kind.value === 'formulas')
const byRule = computed(() => kind.value === 'conditionalFormats' || kind.value === 'validation')
const ruled = (k: SpecialKind) => k === 'conditionalFormats' || k === 'validation'
/** Constants or Formulas with no kind of value checked find nothing to look for. */
const canRun = computed(() => !byValue.value || values.value.length > 0)
function ok(): void {
  if (canRun.value) void special.run({ kind: kind.value, values: SPECIAL_VALUES.filter((v) => values.value.includes(v)), same: same.value })
}
</script>

<template>
  <DialogShell class="xs-dialog xs-dialog-wide" :aria-labelledby="titleId" focus="[data-xs-special-kind]:checked" data-xs-special @close="special.close()">
    <p :id="titleId" class="xs-modal-title">{{ m.specialTitle }}</p>
    <div class="xs-dialog-groups">
      <fieldset class="xs-dialog-choices" @keydown.enter.prevent="ok()">
        <legend>{{ m.specialSelect }}</legend>
        <div class="xs-dialog-cols">
          <div>
            <label v-for="o in LEFT" :key="o.kind" class="xs-dialog-choice"><input v-model="kind" type="radio" :value="o.kind" :data-xs-special-kind="o.kind" />{{ text(o.label) }}</label>
            <label v-for="v in SPECIAL_VALUES" :key="v" class="xs-dialog-choice xs-dialog-sub"><input v-model="values" type="checkbox" :value="v" :disabled="!byValue" :data-xs-special-value="v" />{{ text(VALUE_LABELS[v]) }}</label>
            <label class="xs-dialog-choice"><input v-model="kind" type="radio" value="blanks" data-xs-special-kind="blanks" />{{ m.specialBlanks }}</label>
            <label class="xs-dialog-choice"><input v-model="kind" type="radio" value="currentRegion" data-xs-special-kind="currentRegion" />{{ m.specialRegion }}</label>
          </div>
          <div>
            <label v-for="o in RIGHT" :key="o.kind" class="xs-dialog-choice"><input v-model="kind" type="radio" :value="o.kind" :disabled="ruled(o.kind) && !special.hasRules.value" :data-xs-special-kind="o.kind" />{{ text(o.label) }}</label>
            <label class="xs-dialog-choice xs-dialog-sub"><input v-model="same" type="radio" :value="false" :disabled="!byRule" data-xs-special-same="all" />{{ m.specialAll }}</label>
            <label class="xs-dialog-choice xs-dialog-sub"><input v-model="same" type="radio" :value="true" :disabled="!byRule" data-xs-special-same="same" />{{ m.specialSame }}</label>
          </div>
        </div>
      </fieldset>
    </div>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button xs-button-main" :disabled="!canRun" data-xs-dialog-ok @click="ok()">{{ m.ok }}</button>
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="special.close()">{{ m.cancel }}</button>
    </div>
  </DialogShell>
</template>
