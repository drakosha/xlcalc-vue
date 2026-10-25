<script setup lang="ts">
// Excel's Zoom dialog (View > Zoom, or a click on the status bar's percentage): 200%, 100%, 75%, 50%, 25%, Fit selection
// and Custom, from 10% to 400%; it opens on the choice matching the sheet's zoom, otherwise on Custom with it.
import { ref, useId } from 'vue'
import { MAX_ZOOM, MIN_ZOOM } from '../core/zoom'
import { useSheetContext } from './context'
import DialogShell from './DialogShell.vue'

const PRESETS = [200, 100, 75, 50, 25]

const ctx = useSheetContext()
const m = ctx.messages
const zm = ctx.zoomer
const now = zm.percent.value
const choice = ref<string>(PRESETS.includes(now) ? String(now) : 'custom')
const custom = ref(String(now))
const error = ref('')
const titleId = `${useId()}-title`

function close(): void {
  zm.dialogOpen.value = false
  ctx.focusGrid()
}
function apply(): void {
  if (choice.value === 'fit') {
    zm.fitSelection(true)
    return close()
  }
  const pct = choice.value === 'custom' ? Number(custom.value.trim().replace(',', '.').replace(/%$/, '')) : Number(choice.value)
  if (!Number.isFinite(pct) || pct < MIN_ZOOM || pct > MAX_ZOOM) {
    error.value = m.value.zoomRange
    return
  }
  zm.choose(pct)
  close()
}
</script>

<template>
  <DialogShell class="xs-dialog" :aria-labelledby="titleId" focus="[data-xs-zoom-choice]:checked" data-xs-zoom-dialog @close="close">
    <p :id="titleId" class="xs-modal-title">{{ m.zoomTitle }}</p>
    <div class="xs-dialog-choices" role="radiogroup" @keydown.enter.prevent="apply">
      <label v-for="p in PRESETS" :key="p" class="xs-dialog-choice"><input v-model="choice" type="radio" :value="String(p)" :data-xs-zoom-choice="p" />{{ p }}%</label>
      <label class="xs-dialog-choice"><input v-model="choice" type="radio" value="fit" data-xs-zoom-choice="fit" />{{ m.zoomFit }}</label>
      <label class="xs-dialog-choice">
        <input v-model="choice" type="radio" value="custom" data-xs-zoom-choice="custom" />{{ m.zoomCustom }}
        <input
          v-model="custom"
          class="xs-field xs-field-s"
          type="text"
          inputmode="numeric"
          autocomplete="off"
          spellcheck="false"
          :aria-label="m.zoomCustom"
          data-xs-zoom-custom
          @focus="choice = 'custom'"
        />%
      </label>
    </div>
    <p v-if="error" class="xs-dialog-error" role="alert">{{ error }}</p>
    <div class="xs-dialog-foot">
      <button type="button" class="xs-button" data-xs-dialog-cancel @click="close">{{ m.cancel }}</button>
      <button type="button" class="xs-button xs-button-main" data-xs-dialog-ok @click="apply">{{ m.ok }}</button>
    </div>
  </DialogShell>
</template>
