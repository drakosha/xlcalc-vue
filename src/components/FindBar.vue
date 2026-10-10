<script setup lang="ts">
// Search on the sheet under the formula bar: the source finds, hits in the frame are tinted, Enter steps through them.
import { useSheetContext } from './context'
import { IconClose, IconDown, IconUp } from './icons'

const ctx = useSheetContext()
const { find } = ctx
const m = ctx.messages
const setInput = (el: unknown) => (find.input.value = (el as HTMLInputElement | null) ?? null)
</script>

<template>
  <div class="xs-find" role="search">
    <input
      :ref="setInput"
      v-model="find.text.value"
      class="xs-find-input"
      type="search"
      spellcheck="false"
      :placeholder="m.findPlaceholder"
      @keydown="find.onFindKey"
    />
    <span class="xs-find-status" aria-live="polite">{{ find.status.value }}</span>
    <button type="button" class="xs-tool" :title="m.findPrev" :disabled="!find.hits.value.length" @click="find.step(-1)">
      <IconUp />
    </button>
    <button type="button" class="xs-tool" :title="m.findNext" :disabled="!find.hits.value.length" @click="find.step(1)">
      <IconDown />
    </button>
    <button type="button" class="xs-tool" :title="m.close" @click="find.closeFind()">
      <IconClose />
    </button>
  </div>
</template>
