<script setup lang="ts">
// Excel's dialogs for defined names: the Name Manager (its list of Name, Refers To and Scope, New, Edit and Delete, and Refers
// to edited in the box at the bottom with Commit and Cancel), the New Name and Edit Name dialog, and the alert over them for a
// refusal and Excel's questions. The state and the calls to the source are `useNameDialogs`' (src/core/name-dialogs.ts).
import { computed, nextTick, ref, useId, watch } from 'vue'
import type { SheetName } from '../contract/workbook'
import { nameKey } from '../core/names'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import { IconCheck, IconClose } from './icons'
import '../controls.css'
import '../names.css'

const ctx = useSheetContext()
const nd = ctx.nameDialogs
const m = ctx.messages
const id = useId()

const manager = nd.manager
const draft = nd.draft
const ask = nd.ask
const editable = ctx.names.editable
/** Scope offers the workbook and its sheets, except very hidden ones, which Excel never shows. */
const scopes = computed(() => ctx.sheets.value.filter((x) => x.hidden !== 'very').map((x) => x.name))
const scopeText = (n: SheetName): string => n.scope ?? m.value.nameWorkbook

const list = ref<HTMLElement | null>(null)
/** Down and Up move the selection in the list, Home and End to its ends; Enter edits the selected name, as a double click does. */
function onListKey(e: KeyboardEvent): void {
  const rows = nd.rows.value
  if (!rows.length) return
  const at = rows.findIndex((n) => nameKey(n) === manager.value?.selected)
  let next: number
  if (e.key === 'ArrowDown') next = Math.min(rows.length - 1, at + 1)
  else if (e.key === 'ArrowUp') next = Math.max(0, at - 1)
  else if (e.key === 'Home') next = 0
  else if (e.key === 'End') next = rows.length - 1
  else if (e.key === 'Enter' && at >= 0) {
    e.preventDefault()
    return nd.edit()
  } else return
  e.preventDefault()
  nd.select(rows[next]!)
  void nextTick(() => list.value?.querySelector('[aria-selected=true]')?.scrollIntoView?.({ block: 'nearest' }))
}
// Back from New Name or Edit Name, the Name Manager's list takes focus again.
watch(draft, (d, was) => {
  if (!d && was && manager.value) void nextTick(() => list.value?.focus())
})
const onRefInput = (e: Event) => nd.editRef((e.target as HTMLInputElement).value)
/** Esc in Refers to cancels its edit; with nothing edited it closes the Name Manager, as Esc does elsewhere in it. */
function onRefEsc(e: KeyboardEvent): void {
  if (!nd.refChanged.value) return
  e.preventDefault()
  e.stopPropagation()
  nd.cancelRef()
}
</script>

<template>
  <DialogShell v-if="manager" class="xs-dialog xs-nm" full focus="[data-xs-names-list]" :aria-labelledby="`${id}-nm`" data-xs-name-manager @close="nd.closeManager()">
    <p :id="`${id}-nm`" class="xs-modal-title">{{ m.nameManager }}</p>
    <div v-if="editable" class="xs-nm-tools">
      <button type="button" class="xs-button" data-xs-name-new @click="nd.define()">{{ m.nameNew }}</button>
      <button type="button" class="xs-button" :disabled="!nd.selected.value" data-xs-name-edit @click="nd.edit()">{{ m.nameEdit }}</button>
      <button type="button" class="xs-button" :disabled="!nd.selected.value" data-xs-name-delete @click="nd.remove()">{{ m.nameDelete }}</button>
    </div>
    <div ref="list" class="xs-nm-list" role="listbox" tabindex="0" :aria-label="m.nameManager" data-xs-names-list @keydown="onListKey">
      <div class="xs-nm-row xs-nm-head" aria-hidden="true">
        <span>{{ m.nameColumn('name') }}</span>
        <span>{{ m.nameColumn('refersTo') }}</span>
        <span>{{ m.nameColumn('scope') }}</span>
      </div>
      <div
        v-for="n in nd.rows.value"
        :key="nameKey(n)"
        :class="['xs-nm-row', { 'xs-nm-on': nameKey(n) === manager.selected }]"
        role="option"
        :aria-selected="nameKey(n) === manager.selected"
        :data-xs-name-row="n.name"
        @click="nd.select(n)"
        @dblclick="nd.edit()"
      >
        <span :title="n.name">{{ n.name }}</span>
        <span :title="`=${n.ref}`">={{ n.ref }}</span>
        <span>{{ scopeText(n) }}</span>
      </div>
    </div>
    <div class="xs-dialog-field">
      <label :for="`${id}-ref`">{{ m.nameRefersTo }}</label>
      <span class="xs-nm-ref">
        <button v-if="editable" type="button" class="xs-tool" :disabled="!nd.refChanged.value" :title="m.cancel" :aria-label="m.cancel" data-xs-name-ref-cancel @click="nd.cancelRef()">
          <IconClose />
        </button>
        <button v-if="editable" type="button" class="xs-tool" :disabled="!nd.refChanged.value" :title="m.nameCommit" :aria-label="m.nameCommit" data-xs-name-ref-commit @click="nd.commitRef()">
          <IconCheck />
        </button>
        <input
          :id="`${id}-ref`"
          class="xs-field xs-field-l"
          type="text"
          spellcheck="false"
          autocomplete="off"
          :value="nd.refText.value"
          :readonly="!editable || !nd.selected.value"
          data-xs-name-ref
          @input="onRefInput"
          @keydown.enter.prevent="nd.commitRef()"
          @keydown.esc="onRefEsc"
        />
      </span>
    </div>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-name-close @click="nd.closeManager()">{{ m.close }}</button>
    </div>
  </DialogShell>
  <DialogShell
    v-if="draft"
    as="form"
    class="xs-dialog xs-nd"
    full
    focus="[data-xs-name-field]"
    :aria-labelledby="`${id}-nd`"
    data-xs-name-dialog
    @submit.prevent="nd.submit()"
    @close="nd.closeDraft()"
  >
    <p :id="`${id}-nd`" class="xs-modal-title">{{ draft.editing === null ? m.nameNewTitle : m.nameEditTitle }}</p>
    <label class="xs-dialog-field">
      <span>{{ m.nameLabel }}</span>
      <input v-model="draft.name" class="xs-field xs-field-l" type="text" spellcheck="false" autocomplete="off" data-xs-name-field />
    </label>
    <label class="xs-dialog-field">
      <span>{{ m.nameScope }}</span>
      <!-- Excel's Edit Name does not change a name's scope -->
      <select v-model="draft.scope" class="xs-field xs-field-l" :disabled="draft.editing !== null" data-xs-name-scope>
        <option value="">{{ m.nameWorkbook }}</option>
        <option v-for="sheet in scopes" :key="sheet" :value="sheet">{{ sheet }}</option>
      </select>
    </label>
    <label class="xs-dialog-field">
      <span>{{ m.nameRefersTo }}</span>
      <input v-model="draft.ref" class="xs-field xs-field-l" type="text" spellcheck="false" autocomplete="off" data-xs-name-refers />
    </label>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-name-cancel @click="nd.closeDraft()">{{ m.cancel }}</button>
      <button type="submit" class="xs-button xs-button-main" :disabled="nd.busy.value" data-xs-name-ok>{{ m.ok }}</button>
    </div>
  </DialogShell>
  <DialogShell
    v-if="ask"
    class="xs-dialog"
    role="alertdialog"
    :outside="false"
    :focus="`[data-xs-name-answer='${ask.kind === 'save' ? 'yes' : 'ok'}']`"
    :aria-labelledby="`${id}-alert`"
    :aria-describedby="`${id}-text`"
    :data-xs-name-alert="ask.kind"
    @keydown.stop
    @close="nd.answer(false)"
  >
    <p :id="`${id}-alert`" class="xs-modal-title">{{ m.validationTitle }}</p>
    <p :id="`${id}-text`">{{ ask.message }}</p>
    <div class="xs-dialog-foot">
      <template v-if="ask.kind === 'save'">
        <button type="button" class="xs-button xs-button-main" data-xs-name-answer="yes" @click="nd.answer(true)">{{ m.validationYes }}</button>
        <button type="button" class="xs-button" data-xs-name-answer="no" @click="nd.answer(false)">{{ m.validationNo }}</button>
      </template>
      <template v-else>
        <button type="button" class="xs-button xs-button-main" data-xs-name-answer="ok" @click="nd.answer(true)">{{ m.ok }}</button>
        <button v-if="ask.kind === 'delete'" type="button" class="xs-button" data-xs-name-answer="cancel" @click="nd.answer(false)">{{ m.cancel }}</button>
      </template>
    </div>
  </DialogShell>
</template>
