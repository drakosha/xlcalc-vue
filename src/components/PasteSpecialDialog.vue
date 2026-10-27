<script setup lang="ts">
// Excel's Paste Special dialog (Ctrl+Alt+V, the cell menu) for the copied range: Paste All, Formulas, Values, Formats or
// Validation, and Transpose; OK pastes at the selection as a single undo step. The choices xlcalc's copyRange takes.
import { ref, useId } from 'vue'
import type { PasteKind } from 'xlcalc/contract'
import type { SheetMessages } from '../contract/messages'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'

const KINDS: { kind: PasteKind; label: keyof SheetMessages }[] = [
  { kind: 'all', label: 'pasteSpecialAll' },
  { kind: 'formulas', label: 'pasteSpecialFormulas' },
  { kind: 'values', label: 'pasteSpecialValues' },
  { kind: 'formats', label: 'pasteSpecialFormats' },
  { kind: 'validation', label: 'pasteSpecialValidation' },
]

const ctx = useSheetContext()
const m = ctx.messages
const kind = ref<PasteKind>('all')
const transpose = ref(false)
const titleId = `${useId()}-title`
const text = (key: keyof SheetMessages) => m.value[key] as string
</script>

<template>
  <DialogShell class="xs-dialog" :aria-labelledby="titleId" focus="[data-xs-paste-kind]:checked" data-xs-paste-special @close="ctx.clip.closeSpecial()">
    <p :id="titleId" class="xs-modal-title">{{ m.pasteSpecialTitle }}</p>
    <fieldset class="xs-dialog-choices" @keydown.enter.prevent="ctx.clip.closeSpecial({ paste: kind, transpose })">
      <legend>{{ m.pasteSpecialPaste }}</legend>
      <label v-for="k in KINDS" :key="k.kind" class="xs-dialog-choice"><input v-model="kind" type="radio" :value="k.kind" :data-xs-paste-kind="k.kind" />{{ text(k.label) }}</label>
      <label class="xs-dialog-choice"><input v-model="transpose" type="checkbox" data-xs-paste-transpose />{{ m.pasteSpecialTranspose }}</label>
    </fieldset>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="ctx.clip.closeSpecial()">{{ m.cancel }}</button>
      <button type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="ctx.clip.closeSpecial({ paste: kind, transpose })">{{ m.ok }}</button>
    </div>
  </DialogShell>
</template>
