<script setup lang="ts">
// A colour of the Format Cells dialog: Automatic or No Color, Excel's palette as on the toolbar, and a #RRGGBB field.
import { ref, watch } from 'vue'
import type { SheetMessages } from '../contract/messages'
import ColourSwatches from './ColourSwatches.vue'

const props = defineProps<{
  /** `#RRGGBB`, or null for the choice without a colour. */
  modelValue: string | null
  /** The choice without a colour: Automatic for text and lines, No Color for a fill. */
  noneLabel: string
  messages: SheetMessages
  /** Names the picker's buttons for tests and checks (`data-fmt-colour-of`). */
  name: string
}>()
const emit = defineEmits<{
  'update:modelValue': [colour: string | null]
  /** The field holds what is no colour: OK waits until it is one. */
  invalid: [bad: boolean]
}>()

const HEX = /^#?([0-9a-f]{6})$/i
/** A code as `#RRGGBB`, null for what is no colour. */
const hexOf = (text: string): string | null => {
  const m = HEX.exec(text.trim())
  return m ? `#${m[1]!.toUpperCase()}` : null
}
const code = ref(props.modelValue ?? '')
const bad = ref(false)
watch(
  () => props.modelValue,
  (c) => {
    if (c !== hexOf(code.value)) code.value = c ?? ''
    setBad(false)
  },
)
function setBad(value: boolean): void {
  if (bad.value === value) return
  bad.value = value
  emit('invalid', value)
}
function pick(c: string | null): void {
  code.value = c ?? ''
  setBad(false)
  emit('update:modelValue', c)
}
/** A code typed: a colour once it is six hex digits, with or without `#`; empty is the choice without a colour. */
function typed(): void {
  const hex = hexOf(code.value)
  if (hex) {
    setBad(false)
    emit('update:modelValue', hex)
  } else if (!code.value.trim()) pick(null)
  else setBad(true)
}
</script>

<template>
  <div class="xs-fmt-colour" :data-fmt-colour-of="name">
    <button type="button" class="xs-fmt-none" :aria-pressed="!modelValue" data-fmt-colour-none @click="pick(null)">{{ noneLabel }}</button>
    <ColourSwatches :model-value="modelValue" :label="messages.fmtColor" @update:model-value="pick" />
    <label class="xs-fmt-field">
      <span>{{ messages.fmtColorCode }}</span>
      <input v-model="code" class="xs-field xs-fmt-input xs-fmt-hex" type="text" maxlength="7" spellcheck="false" autocomplete="off" placeholder="#RRGGBB" :aria-invalid="bad" data-fmt-colour-code @input="typed" />
    </label>
    <p v-if="bad" class="xs-fmt-error" role="alert">{{ messages.fmtColorBad }}</p>
  </div>
</template>
