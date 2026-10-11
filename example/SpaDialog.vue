<script setup lang="ts">
// A dialog over the app: a title, its fields and its buttons; Esc and a press on the backdrop close it.
import { onMounted, ref } from 'vue'

defineProps<{ title: string }>()
const emit = defineEmits<{ close: [] }>()
const box = ref<HTMLElement | null>(null)
// The first field takes the focus, so Esc and Enter work at once.
onMounted(() => box.value?.querySelector<HTMLElement>('input, select, button')?.focus())
</script>

<template>
  <div class="spa-backdrop" @pointerdown.self="emit('close')" @keydown.esc="emit('close')">
    <form ref="box" class="spa-dialog" role="dialog" aria-modal="true" :aria-label="title" @submit.prevent>
      <h2 class="spa-dialog-title">{{ title }}</h2>
      <slot />
      <div class="spa-dialog-foot"><slot name="foot" /></div>
    </form>
  </div>
</template>
