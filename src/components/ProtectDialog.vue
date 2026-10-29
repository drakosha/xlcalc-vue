<script setup lang="ts">
// Excel's Protect Sheet dialog without a password, which the workbook does not hold: Protect the sheet and contents of
// locked cells (OK is disabled without it), then what users of the sheet may still do, in Excel's order, with its defaults;
// Select locked cells takes Select unlocked cells along, as in Excel.
import { useId } from 'vue'
import { PROTECTION_BOXES, type ProtectionAllowance } from '../protection'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'

const ctx = useSheetContext()
const m = ctx.messages
const p = ctx.protection
const titleId = `${useId()}-title`

function toggle(box: ProtectionAllowance, on: boolean): void {
  const boxes = p.draft.value?.boxes
  if (!boxes) return
  boxes[box] = on
  if (box === 'selectLockedCells' && on) boxes.selectUnlockedCells = true
  if (box === 'selectUnlockedCells' && !on) boxes.selectLockedCells = false
}
</script>

<template>
  <DialogShell v-if="p.draft.value" class="xs-dialog" :aria-labelledby="titleId" focus="[data-xs-protect-ok]" data-xs-protect-dialog @close="p.close()">
    <p :id="titleId" class="xs-modal-title">{{ m.protectTitle }}</p>
    <label class="xs-dialog-choice"><input v-model="p.draft.value.contents" type="checkbox" data-xs-protect-contents />{{ m.protectContents }}</label>
    <fieldset class="xs-dialog-choices">
      <legend>{{ m.protectAllow }}</legend>
      <div class="xs-dialog-list xs-protect-list">
        <label v-for="box in PROTECTION_BOXES" :key="box" class="xs-dialog-choice">
          <input type="checkbox" :checked="p.draft.value.boxes[box]" :data-xs-protect-box="box" @change="toggle(box, ($event.target as HTMLInputElement).checked)" />{{ m.protectBox(box) }}
        </label>
      </div>
    </fieldset>
    <p class="xs-muted">{{ m.protectNote }}</p>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button xs-button-main" :disabled="!p.draft.value.contents" data-xs-protect-ok data-xs-dialog-ok @click="p.apply()">{{ m.ok }}</button>
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="p.close()">{{ m.cancel }}</button>
    </div>
  </DialogShell>
</template>

