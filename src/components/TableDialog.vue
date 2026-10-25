<script setup lang="ts">
// Excel's Create Table dialog (the range, My table has headers, OK and Cancel) and Convert to Range's question (Yes, No).
// Enter in the range box is OK; Esc or a click outside is Cancel or No.
import { computed, useId } from 'vue'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import '../controls.css'

const ctx = useSheetContext()
const { tables } = ctx
const m = ctx.messages
const titleId = `${useId()}-title`

const create = computed(() => (tables.dialog.value?.kind === 'create' ? tables.dialog.value : null))
const convert = computed(() => (tables.dialog.value?.kind === 'convert' ? tables.dialog.value : null))
</script>

<template>
  <DialogShell
    v-if="tables.dialog.value"
    class="xs-dialog"
    :focus="create ? '[data-xs-table-range]' : '[data-xs-dialog-ok]'"
    :aria-labelledby="titleId"
    data-xs-table-dialog
    @close="tables.close()"
  >
    <template v-if="create">
      <p :id="titleId" class="xs-modal-title">{{ m.tableCreate }}</p>
      <label class="xs-dialog-field">
        <span>{{ m.tableWhere }}</span>
        <input
          v-model="create.draft"
          class="xs-field xs-field-l"
          type="text"
          autocomplete="off"
          spellcheck="false"
          data-xs-table-range
          @input="create.error = ''"
          @keydown.enter.prevent="tables.create()"
        />
      </label>
      <p v-if="create.error" class="xs-dialog-error" role="alert">{{ create.error }}</p>
      <label class="xs-dialog-choice">
        <input v-model="create.headers" type="checkbox" :disabled="!tables.canInsertHeader.value" data-xs-table-headers />{{ m.tableHeaders }}
      </label>
      <div class="xs-dialog-foot">
        <button type="button" class="xs-button" data-xs-dialog-cancel @click="tables.close()">{{ m.cancel }}</button>
        <button type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="tables.create()">{{ m.ok }}</button>
      </div>
    </template>
    <template v-else-if="convert">
      <p :id="titleId" class="xs-modal-title">{{ m.validationTitle }}</p>
      <p>{{ m.tableConvert }}</p>
      <div class="xs-dialog-foot">
        <button type="button" class="xs-button" data-xs-dialog-cancel @click="tables.close(false)">{{ m.validationNo }}</button>
        <button type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="tables.close(true)">{{ m.validationYes }}</button>
      </div>
    </template>
  </DialogShell>
</template>
