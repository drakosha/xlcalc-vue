<script setup lang="ts">
// A color in the Format Cells dialog: Automatic or No Color, Excel's palette as on the toolbar, and a #RRGGBB input.
import { ref, watch } from 'vue'
import type { SheetMessages } from '../contract/messages'
import ColourSwatches from './ColourSwatches.vue'

const props = defineProps<{
  /** `#RRGGBB`, or null for the no-color choice. */
  modelValue: string | null
  /** The no-color choice: Automatic for text and lines, No Color for a fill. */
  noneLabel: string
  messages: SheetMessages
  /** Names the picker's buttons for tests and checks (`data-fmt-colour-of`). */
  name: string
}>()
const emit = defineEmits<{
  'update:modelValue': [colour: string | null]
  /** The input holds an invalid color: OK waits until it is valid. */
  invalid: [bad: boolean]
}>()

const HEX = /^#?([0-9a-f]{6})$/i
/** A code as `#RRGGBB`, null for an invalid color. */
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
/** A typed code: a color once it is six hex digits, with or without `#`; empty means the no-color choice. */
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
