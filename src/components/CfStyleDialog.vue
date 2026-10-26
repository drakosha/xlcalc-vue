<script setup lang="ts">
// Format Cells for a conditional format (Custom Format, and Format… in the rule dialogs): the Font, Border and Fill tabs of
// Excel's dialog for a rule, which sets a font style, an underline, strikethrough and a color, the cell's outline and a fill;
// a choice left as it opened is not part of the rule. Clear empties every tab.
import { reactive, ref, shallowRef } from 'vue'
import type { BorderLine, ConditionalStyle } from 'xlcalc/contract'
import type { SheetMessages } from '../contract/messages'
import { FONT_STYLES, startOf, UNDERLINES, type FormatStart } from '../format-cells'
import DialogShell from './DialogShell.vue'
import FormatBorders from './FormatBorders.vue'
import FormatColour from './FormatColour.vue'
import { IconClose } from './icons'

const props = defineProps<{ style: ConditionalStyle; messages: SheetMessages }>()
const emit = defineEmits<{ done: [style: ConditionalStyle]; close: [] }>()
const m = props.messages

const TABS = ['font', 'border', 'fill'] as const
const SIDES = ['top', 'right', 'bottom', 'left'] as const
const tab = ref<(typeof TABS)[number]>('font')

type FontStyleKind = (typeof FONT_STYLES)[number] | ''
const f = props.style.font
const fontStyle = ref<FontStyleKind>(f?.bold ? (f.italic ? 'boldItalic' : 'bold') : f?.italic ? 'italic' : f?.bold === false && f.italic === false ? 'regular' : '')
const underline = ref<string>(f?.underline ?? 'none')
const strike = ref(!!f?.strike)
const fontColour = ref<string | null>(f?.color ?? null)
const fill = ref<string | null>(props.style.fill?.color ?? null)
/** The Border tab over one cell: its outline is the rule's border, and Inside is off, as for a rule in Excel. */
const start = shallowRef<FormatStart>(edgesStart(props.style.border))
const draft = ref(structuredClone(start.value.draft))
const badColour = reactive<Record<string, boolean>>({})

function edgesStart(border: ConditionalStyle['border']): FormatStart {
  const s = startOf({ sheet: '', ref: 'A1' })
  return { ...s, edges: Object.fromEntries(SIDES.map((k) => [k, border?.[k] ?? null])) }
}

function clear(): void {
  fontStyle.value = ''
  underline.value = 'none'
  strike.value = false
  fontColour.value = null
  fill.value = null
  start.value = edgesStart(undefined)
  draft.value = structuredClone(start.value.draft)
}

function ok(): void {
  if (Object.values(badColour).some(Boolean)) return
  const font: NonNullable<ConditionalStyle['font']> = {}
  const k = fontStyle.value
  if (k === 'regular') Object.assign(font, { bold: false, italic: false })
  if (k === 'bold' || k === 'boldItalic') font.bold = true
  if (k === 'italic' || k === 'boldItalic') font.italic = true
  if (underline.value === 'single' || underline.value === 'double') font.underline = underline.value
  if (strike.value) font.strike = true
  if (fontColour.value) font.color = fontColour.value
  const border: Record<string, BorderLine> = {}
  for (const side of SIDES) {
    const line = draft.value.edges[side] !== undefined ? draft.value.edges[side] : start.value.edges[side]
    if (line) border[side] = { ...line }
  }
  const out: ConditionalStyle = {
    ...(Object.keys(font).length ? { font } : {}),
    ...(fill.value ? { fill: { color: fill.value } } : {}),
    ...(Object.keys(border).length ? { border } : {}),
    ...(props.style.numberFormat !== undefined ? { numberFormat: props.style.numberFormat } : {}),
  }
  emit('done', out)
}
</script>

<template>
  <DialogShell as="form" class="xs-fmt" :outside="false" full focus="[aria-selected='true']" :aria-label="m.formatCells" data-cf-style novalidate @submit.prevent="ok" @close="emit('close')">
    <header class="xs-fmt-head">
      <span class="xs-modal-title xs-fmt-title">{{ m.formatCells }}</span>
      <button type="button" class="xs-tool xs-fmt-x" :title="m.cancel" :aria-label="m.cancel" @click="emit('close')">
        <IconClose />
      </button>
    </header>
    <div class="xs-fmt-tabs" role="tablist">
      <button v-for="t in TABS" :key="t" type="button" class="xs-fmt-tab" role="tab" :aria-selected="tab === t" :data-fmt-tab="t" @click="tab = t">
        {{ m.fmtTab(t) }}
      </button>
    </div>
    <div class="xs-fmt-body" role="tabpanel" :aria-label="m.fmtTab(tab)">
      <div v-show="tab === 'font'" class="xs-fmt-grid">
        <label class="xs-fmt-field">
          <span>{{ m.fmtFontStyle }}</span>
          <select v-model="fontStyle" class="xs-field xs-fmt-select" data-fmt-font-style>
            <option value="">{{ m.cfNotSet }}</option>
            <option v-for="k in FONT_STYLES" :key="k" :value="k">{{ m.fmtFontStyleKind(k) }}</option>
          </select>
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.fmtUnderline }}</span>
          <select v-model="underline" class="xs-field xs-fmt-select" data-fmt-underline>
            <option v-for="u in UNDERLINES" :key="u" :value="u">{{ m.fmtUnderlineKind(u) }}</option>
          </select>
        </label>
        <label class="xs-fmt-check">
          <input v-model="strike" type="checkbox" data-fmt-strike />
          {{ m.fmtStrike }}
        </label>
        <fieldset class="xs-fmt-box xs-fmt-wide">
          <legend>{{ m.fmtColor }}</legend>
          <FormatColour v-model="fontColour" :none-label="m.colorAuto" :messages="m" name="font" @invalid="badColour.font = $event" />
        </fieldset>
      </div>
      <FormatBorders v-show="tab === 'border'" :key="JSON.stringify(start.edges)" :start="start" :draft="draft" :messages="m" @invalid="badColour.line = $event" />
      <fieldset v-show="tab === 'fill'" class="xs-fmt-box">
        <legend>{{ m.fmtBackground }}</legend>
        <FormatColour v-model="fill" :none-label="m.fillNone" :messages="m" name="fill" @invalid="badColour.fill = $event" />
      </fieldset>
    </div>
    <footer class="xs-fmt-foot">
      <button type="button" class="xs-button xs-fmt-btn xs-cf-clear" data-cf-clear @click="clear">{{ m.cfClear }}</button>
      <button type="button" class="xs-button xs-fmt-btn" data-fmt-cancel @click="emit('close')">{{ m.cancel }}</button>
      <button type="submit" class="xs-button xs-button-main xs-fmt-btn" data-fmt-ok>{{ m.ok }}</button>
    </footer>
  </DialogShell>
</template>

<style src="../format-dialog.css"></style>
<style src="../cf.css"></style>
