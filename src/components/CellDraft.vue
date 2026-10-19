<script setup lang="ts">
// The draft typed in the formula bar, displayed in its cell: a separate component, so a typed character redraws this text and not the grid.
import { computed } from 'vue'
import { fillText } from '../core/cell-style'
import { useSheetContext } from './context'

/** `fill`: the cell's width when its alignment repeats the text across it (Excel's Fill). */
const props = defineProps<{ fill?: number }>()
const ctx = useSheetContext()
const text = computed(() => {
  const t = ctx.editor.value?.text ?? ''
  return props.fill === undefined ? t : fillText(t, props.fill)
})
</script>

<template>{{ text }}</template>
