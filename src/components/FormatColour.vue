<script setup lang="ts">
// A colour of the Format Cells dialog: Automatic or No Color, Excel's palette as on the toolbar, and a #RRGGBB field.
import { ref, watch } from 'vue'
import type { SheetMessages } from '../messages'
import { PALETTE } from '../toolbar'

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
const code = ref(props.modelValue ?? '')
const bad = ref(false)
watch(
  () => props.modelValue,
  (c) => {
    if (c !== (HEX.test(code.value) ? `#${HEX.exec(code.value)![1]!.toUpperCase()}` : null)) code.value = c ?? ''
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
  const t = code.value.trim()
  const m = HEX.exec(t)
  if (m) {
    setBad(false)
    emit('update:modelValue', `#${m[1]!.toUpperCase()}`)
  } else if (!t) pick(null)
  else setBad(true)
}
</script>

<template>
  <div class="xs-fmt-colour" :data-fmt-colour-of="name">
    <button type="button" class="xs-fmt-none" :aria-pressed="!modelValue" data-fmt-colour-none @click="pick(null)">{{ noneLabel }}</button>
    <div class="xs-fmt-swatches" role="group" :aria-label="messages.fmtColour">
      <button
        v-for="(c, i) in PALETTE"
        :key="i"
        type="button"
        class="xs-fmt-swatch"
        :style="{ backgroundColor: c }"
        :title="c"
        :aria-label="c"
        :aria-pressed="modelValue === c"
        :data-fmt-colour="c"
        @click="pick(c)"
      />
    </div>
    <label class="xs-fmt-field">
      <span>{{ messages.fmtColourCode }}</span>
      <input v-model="code" class="xs-fmt-input xs-fmt-hex" type="text" maxlength="7" spellcheck="false" autocomplete="off" placeholder="#RRGGBB" :aria-invalid="bad" data-fmt-colour-code @input="typed" />
    </label>
    <p v-if="bad" class="xs-fmt-error" role="alert">{{ messages.fmtColourBad }}</p>
  </div>
</template>
