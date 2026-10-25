<script setup lang="ts">
// Excel's Go To dialog (F5, Ctrl+G): the list of locations its jumps have left and of the workbook's names, the Reference box, OK and Cancel.
import { ref, useId } from 'vue'
import { stepIndex } from '../core/list-keys'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { keepActiveInView } from './popup'
import '../controls.css'

const ctx = useSheetContext()
const m = ctx.messages
const goto = ctx.goto
const id = useId()

/** The list as it was when the dialog opened; the Reference box starts with its first location, so F5 and Enter go back. */
const items = goto.items.value
const text = ref(goto.places.value[0] ?? '')
const index = ref(-1)
const list = ref<HTMLElement | null>(null)
const input = ref<HTMLInputElement | null>(null)
keepActiveInView(list, index, () => index.value)
let busy = false

/** A click on an item, or an arrow key, puts it into the Reference box, as in Excel. */
function pick(i: number): void {
  index.value = i
  text.value = items[i] ?? text.value
}
async function ok(): Promise<void> {
  if (busy) return
  busy = true
  const went = await goto.go(text.value)
  busy = false
  // the alert for a bad reference has closed: the box takes focus again with its text selected, as Excel's does
  if (!went) {
    input.value?.focus()
    input.value?.select()
  }
}
/** Up and Down move through the list from anywhere in the dialog. */
function onKey(e: KeyboardEvent): void {
  if ((e.key !== 'ArrowDown' && e.key !== 'ArrowUp') || !items.length) return
  e.preventDefault()
  pick(stepIndex(index.value, e.key === 'ArrowDown' ? 1 : -1, items.length))
  input.value?.select()
}
</script>

<template>
  <DialogShell as="form" class="xs-goto" full focus="[data-xs-goto-ref]" :aria-label="m.goToTitle" data-xs-goto novalidate @submit.prevent="ok()" @keydown="onKey" @close="goto.close()">
    <p class="xs-modal-title">{{ m.goToTitle }}</p>
    <p :id="`${id}-list`">{{ m.goToList }}</p>
    <div ref="list" class="xs-goto-list" role="listbox" :aria-labelledby="`${id}-list`">
      <div
        v-for="(item, i) in items"
        :key="item"
        role="option"
        :aria-selected="i === index ? 'true' : 'false'"
        :class="['xs-goto-item', i === index ? 'xs-goto-on' : '']"
        data-xs-goto-item
        @mousedown.prevent
        @click="pick(i)"
        @dblclick="pick(i), ok()"
      >
        {{ item }}
      </div>
    </div>
    <label class="xs-goto-field">
      <span>{{ m.goToReference }}</span>
      <input ref="input" v-model="text" class="xs-field xs-field-l" type="text" enterkeyhint="go" autocomplete="off" autocapitalize="off" spellcheck="false" data-xs-goto-ref />
    </label>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="goto.close()">{{ m.cancel }}</button>
      <button type="submit" class="xs-button xs-button-main" data-xs-dialog-ok>{{ m.ok }}</button>
    </div>
  </DialogShell>
</template>

<style src="../dialogs.css"></style>
