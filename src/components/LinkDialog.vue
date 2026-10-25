<script setup lang="ts">
// Excel's Insert Hyperlink dialog (Edit Hyperlink on a cell with a link): Text to Display, Link to an existing file or web page
// (its address) or a place in this document (a cell reference on a sheet from the list, or a defined name), the ScreenTip,
// and Remove Link when it edits one. Enter is OK, Esc is Cancel.
import { computed, useId } from 'vue'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'
import '../controls.css'

const ctx = useSheetContext()
const { links } = ctx
const m = ctx.messages
const d = computed(() => links.draft.value)
const titleId = `${useId()}-title`

function onText(e: Event): void {
  if (!d.value || d.value.text === null) return
  d.value.text = (e.target as HTMLInputElement).value
  d.value.textTyped = true
}
function onAddress(e: Event, field: 'address' | 'ref'): void {
  if (!d.value) return
  d.value[field] = (e.target as HTMLInputElement).value
  d.value.error = ''
  links.follow()
}
function pick(tab: 'file' | 'place'): void {
  if (!d.value) return
  d.value.tab = tab
  d.value.error = ''
  links.follow()
}
function choose(key: string): void {
  if (!d.value) return
  d.value.place = key
  d.value.error = ''
  links.follow()
}
</script>

<template>
  <DialogShell
    v-if="d"
    as="form"
    class="xs-link"
    full
    :aria-labelledby="titleId"
    :focus="d.text !== null ? '[data-xs-link-text]' : d.tab === 'file' ? '[data-xs-link-address]' : '[data-xs-link-ref]'"
    data-xs-link
    @submit.prevent="links.apply()"
    @close="links.close()"
  >
    <p :id="titleId" class="xs-modal-title">{{ d.old ? m.linkEditTitle : m.linkInsertTitle }}</p>
    <label class="xs-link-field">
      <span>{{ m.linkText }}</span>
      <input
        class="xs-field xs-field-l"
        type="text"
        autocomplete="off"
        :value="d.text ?? m.linkSelection"
        :disabled="d.text === null"
        data-xs-link-text
        @input="onText"
      />
    </label>
    <div class="xs-link-to" role="radiogroup" :aria-label="m.linkTo">
      <span>{{ m.linkTo }}</span>
      <label class="xs-link-choice"><input type="radio" name="xs-link-to" :checked="d.tab === 'file'" data-xs-link-to="file" @change="pick('file')" />{{ m.linkToFile }}</label>
      <label class="xs-link-choice"><input type="radio" name="xs-link-to" :checked="d.tab === 'place'" data-xs-link-to="place" @change="pick('place')" />{{ m.linkToPlace }}</label>
    </div>
    <label v-if="d.tab === 'file'" class="xs-link-field">
      <span>{{ m.linkAddress }}</span>
      <input class="xs-field xs-field-l" type="text" inputmode="url" autocomplete="off" spellcheck="false" :value="d.address" data-xs-link-address @input="onAddress($event, 'address')" />
    </label>
    <template v-else>
      <label class="xs-link-field">
        <span>{{ m.linkRef }}</span>
        <input class="xs-field xs-field-l" type="text" autocomplete="off" spellcheck="false" :value="d.ref" :disabled="!links.refOn.value" data-xs-link-ref @input="onAddress($event, 'ref')" />
      </label>
      <div class="xs-link-field">
        <span>{{ m.linkPlaces }}</span>
        <div class="xs-link-places" role="listbox" :aria-label="m.linkPlaces">
          <template v-for="g in links.places.value" :key="g.kind">
            <div class="xs-link-group" role="presentation">{{ g.label }}</div>
            <button
              v-for="p in g.places"
              :key="p.key"
              type="button"
              role="option"
              :aria-selected="d.place === p.key"
              :class="['xs-link-place', { 'xs-on': d.place === p.key }]"
              :data-xs-link-place="p.key"
              @click="choose(p.key)"
            >
              {{ p.label }}
            </button>
          </template>
        </div>
      </div>
    </template>
    <label class="xs-link-field">
      <span>{{ m.linkTip }}</span>
      <input v-model="d.tooltip" class="xs-field xs-field-l" type="text" autocomplete="off" data-xs-link-tip />
    </label>
    <p v-if="d.error" class="xs-link-error" role="alert">{{ d.error }}</p>
    <div class="xs-link-foot">
      <button v-if="d.old" type="button" class="xs-button xs-link-remove" data-xs-link-remove @click="links.remove()">{{ m.linkRemove }}</button>
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="links.close()">{{ m.cancel }}</button>
      <button type="submit" class="xs-button xs-button-main" data-xs-dialog-ok>{{ m.ok }}</button>
    </div>
  </DialogShell>
</template>

<style src="../links-notes.css"></style>
