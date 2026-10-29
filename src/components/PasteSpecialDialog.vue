<script setup lang="ts">
// Excel's Paste Special dialog (Ctrl+Alt+V, the cell menu) for the copied range: its Paste choices in two columns, the
// Operation group, Skip Blanks and Transpose; OK pastes at the selection as a single undo step, as xlcalc's copyRange takes them.
// Excel's All using Source theme (All within one workbook) and Paste Link have no counterpart and are left out.
import { ref, useId } from 'vue'
import type { PasteKind, PasteOperation } from 'xlcalc/contract'
import type { SheetMessages } from '../contract/messages'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'

type Choice<T> = { value: T; label: keyof SheetMessages }
/** Excel's Paste choices, column by column as its dialog lays them out. */
const KINDS: Choice<PasteKind>[][] = [
  [
    { value: 'all', label: 'pasteSpecialAll' },
    { value: 'formulas', label: 'pasteSpecialFormulas' },
    { value: 'values', label: 'pasteSpecialValues' },
    { value: 'formats', label: 'pasteSpecialFormats' },
    { value: 'notes', label: 'pasteSpecialNotes' },
    { value: 'validation', label: 'pasteSpecialValidation' },
  ],
  [
    { value: 'allExceptBorders', label: 'pasteSpecialAllExceptBorders' },
    { value: 'columnWidths', label: 'pasteSpecialColumnWidths' },
    { value: 'formulasAndNumberFormats', label: 'pasteSpecialFormulasNumbers' },
    { value: 'valuesAndNumberFormats', label: 'pasteSpecialValuesNumbers' },
    { value: 'allMergingConditionalFormats', label: 'pasteSpecialMergeConditional' },
  ],
]
/** Excel's Operation group, None first. */
const OPERATIONS: Choice<PasteOperation | 'none'>[][] = [
  [
    { value: 'none', label: 'pasteSpecialNone' },
    { value: 'add', label: 'pasteSpecialAdd' },
    { value: 'subtract', label: 'pasteSpecialSubtract' },
  ],
  [
    { value: 'multiply', label: 'pasteSpecialMultiply' },
    { value: 'divide', label: 'pasteSpecialDivide' },
  ],
]

const ctx = useSheetContext()
const m = ctx.messages
const kind = ref<PasteKind>('all')
const operation = ref<PasteOperation | 'none'>('none')
const skipBlanks = ref(false)
const transpose = ref(false)
const titleId = `${useId()}-title`
const text = (key: keyof SheetMessages) => m.value[key] as string
const ok = () => ctx.clip.closeSpecial({ paste: kind.value, transpose: transpose.value, skipBlanks: skipBlanks.value, ...(operation.value !== 'none' ? { operation: operation.value } : {}) })
</script>

<template>
  <DialogShell class="xs-dialog xs-dialog-wide" :aria-labelledby="titleId" focus="[data-xs-paste-kind]:checked" data-xs-paste-special @close="ctx.clip.closeSpecial()">
    <p :id="titleId" class="xs-modal-title">{{ m.pasteSpecialTitle }}</p>
    <div class="xs-dialog-groups" @keydown.enter.prevent="ok()">
      <fieldset class="xs-dialog-choices">
        <legend>{{ m.pasteSpecialPaste }}</legend>
        <div class="xs-dialog-cols">
          <div v-for="(col, i) in KINDS" :key="i">
            <label v-for="k in col" :key="k.value" class="xs-dialog-choice"><input v-model="kind" type="radio" :value="k.value" :data-xs-paste-kind="k.value" />{{ text(k.label) }}</label>
          </div>
        </div>
      </fieldset>
      <fieldset class="xs-dialog-choices">
        <legend>{{ m.pasteSpecialOperation }}</legend>
        <div class="xs-dialog-cols">
          <div v-for="(col, i) in OPERATIONS" :key="i">
            <label v-for="o in col" :key="o.value" class="xs-dialog-choice"><input v-model="operation" type="radio" :value="o.value" :data-xs-paste-operation="o.value" />{{ text(o.label) }}</label>
          </div>
        </div>
      </fieldset>
      <div class="xs-dialog-cols">
        <div><label class="xs-dialog-choice"><input v-model="skipBlanks" type="checkbox" data-xs-paste-skip />{{ m.pasteSpecialSkipBlanks }}</label></div>
        <div><label class="xs-dialog-choice"><input v-model="transpose" type="checkbox" data-xs-paste-transpose />{{ m.pasteSpecialTranspose }}</label></div>
      </div>
    </div>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="ctx.clip.closeSpecial()">{{ m.cancel }}</button>
      <button type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="ok()">{{ m.ok }}</button>
    </div>
  </DialogShell>
</template>
