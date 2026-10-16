<script setup lang="ts">
// Excel's palette as a grid of swatches, the colour given pressed: the toolbar's colour menus and the Format Cells dialog's colours.
import { PALETTE } from '../toolbar'
import '../controls.css'

defineProps<{
  /** The colour shown pressed, `#RRGGBB`; null for none. */
  modelValue: string | null
  /** A name for the group; none where the menu around names it. */
  label?: string
}>()
const emit = defineEmits<{ 'update:modelValue': [colour: string] }>()
</script>

<template>
  <div class="xs-swatches" :role="label ? 'group' : undefined" :aria-label="label">
    <button
      v-for="(c, i) in PALETTE"
      :key="i"
      type="button"
      class="xs-swatch"
      :style="{ backgroundColor: c }"
      :title="c"
      :aria-label="c"
      :aria-pressed="modelValue === c"
      :data-xs-colour="c"
      @click="emit('update:modelValue', c)"
    />
  </div>
</template>
