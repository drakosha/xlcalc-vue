<script setup lang="ts">
import type { RichRun } from '../core/rich-text'
defineProps<{ text: string; runs?: RichRun[] }>()
</script>

<template>
  <template v-if="runs"><span v-for="(run, i) in runs" :key="i" class="xs-rich-run" :style="run.style">{{ run.text }}</span></template>
  <template v-else>{{ text }}</template>
</template>
