<script setup lang="ts">
// Excel's PivotTable from table or range dialog (Insert > PivotTable): the table or range, New Worksheet or Existing
// Worksheet with its location, OK and Cancel. Enter in a box is OK; Esc or a click outside is Cancel.
import { useId } from 'vue'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import '../controls.css'

const ctx = useSheetContext()
const create = ctx.pivotCreate
const m = ctx.messages
const titleId = `${useId()}-title`
</script>

<template>
  <DialogShell
    v-if="create.draft.value"
    class="xs-dialog xs-dialog-wide"
    focus="[data-xs-pivot-source]"
    :aria-labelledby="titleId"
    data-xs-pivot-dialog
    @close="create.close()"
  >
    <p :id="titleId" class="xs-modal-title">{{ m.pivotCreateTitle }}</p>
    <fieldset class="xs-dialog-choices">
      <legend>{{ m.pivotCreateSource }}</legend>
      <label class="xs-dialog-field">
        <span>{{ m.pivotCreateRange }}</span>
        <input
          v-model="create.draft.value.source"
          class="xs-field xs-field-l"
          type="text"
          autocomplete="off"
          spellcheck="false"
          data-xs-pivot-source
          @input="create.draft.value.error = ''"
          @keydown.enter.prevent="create.ok()"
        />
      </label>
    </fieldset>
    <fieldset class="xs-dialog-choices">
      <legend>{{ m.pivotCreatePlace }}</legend>
      <label class="xs-dialog-choice">
        <input v-model="create.draft.value.place" type="radio" value="new" data-xs-pivot-place="new" />{{ m.pivotNewSheet }}
      </label>
      <label class="xs-dialog-choice">
        <input v-model="create.draft.value.place" type="radio" value="existing" data-xs-pivot-place="existing" />{{ m.pivotExistingSheet }}
      </label>
      <label class="xs-dialog-field">
        <span>{{ m.pivotLocation }}</span>
        <input
          v-model="create.draft.value.location"
          class="xs-field xs-field-l"
          type="text"
          autocomplete="off"
          spellcheck="false"
          :disabled="create.draft.value.place !== 'existing'"
          data-xs-pivot-location
          @input="create.draft.value.error = ''"
          @keydown.enter.prevent="create.ok()"
        />
      </label>
    </fieldset>
    <p v-if="create.draft.value.error" class="xs-dialog-error" role="alert" data-xs-pivot-error>{{ create.draft.value.error }}</p>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="create.ok()">{{ m.ok }}</button>
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="create.close()">{{ m.cancel }}</button>
    </div>
  </DialogShell>
</template>
