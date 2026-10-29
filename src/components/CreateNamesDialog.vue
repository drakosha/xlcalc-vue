<script setup lang="ts">
// Excel's Create Names from Selection dialog: Create names in Top row, Left column, Bottom row, Right column, with the boxes
// Excel checks for the selection; OK makes the names (`useCreateNames`), and is off while no box is checked.
import { computed, ref, useId } from 'vue'
import type { SheetMessages } from '../contract/messages'
import type { Labels } from '../core/create-names'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'

const ctx = useSheetContext()
const m = ctx.messages
const cn = ctx.createNames
const titleId = `${useId()}-title`
const SIDES: { side: keyof Labels; label: keyof SheetMessages }[] = [
  { side: 'top', label: 'createNamesTop' },
  { side: 'left', label: 'createNamesLeft' },
  { side: 'bottom', label: 'createNamesBottom' },
  { side: 'right', label: 'createNamesRight' },
]
const labels = ref<Labels>({ ...cn.draft.value!.labels })
const any = computed(() => Object.values(labels.value).some(Boolean))
const text = (key: keyof SheetMessages) => m.value[key] as string
const ok = () => any.value && void cn.submit({ ...labels.value })
</script>

<template>
  <DialogShell class="xs-dialog" :aria-labelledby="titleId" focus="[data-xs-dialog-ok]" data-xs-create-names @close="cn.close()">
    <p :id="titleId" class="xs-modal-title">{{ m.createNamesTitle }}</p>
    <fieldset class="xs-dialog-choices" @keydown.enter.prevent="ok()">
      <legend>{{ m.createNamesIn }}</legend>
      <label v-for="x in SIDES" :key="x.side" class="xs-dialog-choice"><input v-model="labels[x.side]" type="checkbox" :data-xs-create-names-side="x.side" />{{ text(x.label) }}</label>
    </fieldset>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="cn.close()">{{ m.cancel }}</button>
      <button type="button" class="xs-button xs-button-main" :disabled="!any" data-xs-dialog-ok @click="ok()">{{ m.ok }}</button>
    </div>
  </DialogShell>
</template>
