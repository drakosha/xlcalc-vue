<script setup lang="ts">
// Excel's palette as a grid of swatches, with the given color shown pressed: the toolbar's color menus and the Format Cells dialog's colors.
import { PALETTE } from '../toolbar'
import '../controls.css'

defineProps<{
  /** The color shown pressed, `#RRGGBB`; null for none. */
  modelValue: string | null
  /** A label for the group; none where the surrounding menu labels it. */
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
