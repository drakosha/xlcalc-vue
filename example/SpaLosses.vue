<script setup lang="ts">
// What a file lost, like Excel's message bar after opening: one line with each kind and its count, the details on a
// click (each message with its count, the sheets, example cells and names), and a close button.
import { computed, ref } from 'vue'
import type { LossKind } from './spa-open'
import { IconClose, IconWarn } from './spa-icons'
import type { SpaTexts } from './spa-texts'

const props = defineProps<{ kinds: LossKind[]; t: SpaTexts }>()
const emit = defineEmits<{ close: [] }>()
const open = ref(false)
/** Messages of a kind shown in the details; the rest are counted. */
const MESSAGES = 5

const kindText = (k: LossKind) => `${props.t.lossKind[k.code]} (${props.t.count(k.count)})`
const summary = computed(() => props.kinds.map(kindText).join(', '))
</script>

<template>
  <div class="spa-lost" role="status" data-spa-lost>
    <div class="spa-lost-line">
      <IconWarn class="spa-lost-icon" />
      <strong class="spa-lost-title" data-spa-lost-title>{{ t.lostTitle }}</strong>
      <span class="spa-lost-summary" :title="summary" data-spa-lost-summary>{{ summary }}</span>
      <button type="button" class="spa-btn spa-lost-more" :aria-expanded="open" data-spa-lost-more @click="open = !open">{{ open ? t.lostHide : t.lostDetails }}</button>
      <button type="button" class="spa-btn spa-icon-btn" :title="t.close" :aria-label="t.close" data-spa-lost-close @click="emit('close')"><IconClose /></button>
    </div>
    <div v-if="open" class="spa-lost-details" data-spa-lost-details>
      <section v-for="k in kinds" :key="k.code" class="spa-lost-kind" :data-spa-lost-kind="k.code">
        <h3>{{ kindText(k) }}</h3>
        <ul>
          <li v-for="m in k.messages.slice(0, MESSAGES)" :key="m.text">{{ m.text }}<template v-if="m.count > 1"> ({{ t.count(m.count) }})</template></li>
          <li v-if="k.messages.length > MESSAGES">{{ t.lostOther(k.messages.length - MESSAGES) }}</li>
        </ul>
        <p v-if="k.sheets.length">{{ t.lostSheets(k.sheets) }}</p>
        <p v-if="k.cells.count">{{ t.lostCells(k.cells.count, k.cells.examples) }}</p>
        <p v-if="k.names.count">{{ t.lostNames(k.names.count, k.names.examples) }}</p>
      </section>
    </div>
  </div>
</template>
