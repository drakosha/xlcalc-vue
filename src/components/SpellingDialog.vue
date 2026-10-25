<script setup lang="ts">
// Excel's Spelling dialog: Not in Dictionary (the cell's text with the word marked), Suggestions, Ignore Once, Ignore All, Add to
// Dictionary, Change, Change All and Cancel; then Excel's message that the check is complete, with OK. Up and Down move through
// the suggestions, a double click on one is Change, Esc is Cancel.
import { computed, useId } from 'vue'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import '../controls.css'

const ctx = useSheetContext()
const { spell } = ctx
const m = ctx.messages
const id = useId()

/** The cell's text in three pieces around the misspelt word. */
const pieces = computed(() => {
  const sh = spell.shown.value
  if (!sh) return null
  const { at, word } = sh.hit
  return [sh.text.slice(0, at), sh.text.slice(at, at + word.length), sh.text.slice(at + word.length)] as const
})
const suggestions = computed(() => spell.shown.value?.hit.suggestions ?? [])
const canChange = computed(() => spell.canChange.value && suggestions.value.length > 0 && !spell.busy.value)

function onListKey(e: KeyboardEvent): void {
  const sh = spell.shown.value
  const n = suggestions.value.length
  if (!sh || !n || (e.key !== 'ArrowDown' && e.key !== 'ArrowUp')) return
  e.preventDefault()
  sh.choice = (sh.choice + (e.key === 'ArrowDown' ? 1 : n - 1)) % n
}
</script>

<template>
  <DialogShell v-if="spell.shown.value" class="xs-spell" :focus="suggestions.length ? '[data-xs-spell-list]' : '[data-xs-spell=ignoreOnce]'" :aria-labelledby="`${id}-title`" data-xs-spelling @close="spell.close()">
    <p :id="`${id}-title`" class="xs-modal-title">{{ m.spelling }}</p>
    <div class="xs-spell-body">
      <div class="xs-spell-boxes">
        <span :id="`${id}-word`" class="xs-spell-label">{{ m.spellNotInDictionary }}</span>
        <p class="xs-spell-text" :aria-labelledby="`${id}-word`" data-xs-spell-text>{{ pieces![0] }}<mark class="xs-spell-word">{{ pieces![1] }}</mark>{{ pieces![2] }}</p>
        <span :id="`${id}-list`" class="xs-spell-label">{{ m.spellSuggestions }}</span>
        <div
          v-if="suggestions.length"
          class="xs-spell-list"
          role="listbox"
          tabindex="0"
          :aria-labelledby="`${id}-list`"
          :aria-activedescendant="`${id}-s${spell.shown.value.choice}`"
          data-xs-spell-list
          @keydown="onListKey"
          @keydown.enter.prevent="spell.act('change')"
        >
          <div
            v-for="(s, i) in suggestions"
            :id="`${id}-s${i}`"
            :key="s + i"
            role="option"
            :aria-selected="i === spell.shown.value.choice"
            :class="['xs-spell-option', { 'xs-on': i === spell.shown.value.choice }]"
            :data-xs-spell-option="i"
            @click="spell.shown.value!.choice = i"
            @dblclick="spell.act('change')"
          >
            {{ s }}
          </div>
        </div>
        <p v-else class="xs-spell-list xs-spell-none">{{ m.spellNoSuggestions }}</p>
      </div>
      <div class="xs-spell-buttons">
        <button type="button" class="xs-button" :disabled="spell.busy.value" data-xs-spell="ignoreOnce" @click="spell.act('ignoreOnce')">{{ m.spellIgnoreOnce }}</button>
        <button type="button" class="xs-button" :disabled="spell.busy.value" data-xs-spell="ignoreAll" @click="spell.act('ignoreAll')">{{ m.spellIgnoreAll }}</button>
        <button v-if="spell.canLearn.value" type="button" class="xs-button" :disabled="spell.busy.value" data-xs-spell="add" @click="spell.act('add')">{{ m.spellAdd }}</button>
        <button type="button" class="xs-button xs-button-main xs-spell-gap" :disabled="!canChange" data-xs-spell="change" @click="spell.act('change')">{{ m.spellChange }}</button>
        <button type="button" class="xs-button" :disabled="!canChange" data-xs-spell="changeAll" @click="spell.act('changeAll')">{{ m.spellChangeAll }}</button>
      </div>
    </div>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="spell.close()">{{ m.cancel }}</button>
    </div>
  </DialogShell>
  <DialogShell v-else-if="spell.done.value" class="xs-dialog" role="alertdialog" :outside="false" focus="[data-xs-dialog-ok]" :aria-describedby="`${id}-done`" data-xs-spell-done @close="spell.close()">
    <p :id="`${id}-done`">{{ m.spellDone }}</p>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="spell.close()">{{ m.ok }}</button>
    </div>
  </DialogShell>
</template>

<style src="../spelling.css"></style>
