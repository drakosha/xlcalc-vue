<script setup lang="ts">
// Excel's Format Cells dialog (Ctrl+1) for the selection reported by the sheet's `state` event: Number, Alignment, Font,
// Border and Fill, the settings xlcalc's cell style supports; the Number tab is Excel's Number Format box in the workbook's
// locale (xlcalc's list and categories). It opens with the active cell's formatting; OK emits the patches for `SheetView.format`
// covering what changed (one call, a single undo step) and closes, Cancel closes. Full screen on a phone.
import { computed, reactive, ref, shallowRef, toRaw, watch } from 'vue'
import type { SheetFormatPatch, SheetViewState } from '../contract/view'
import {
  draftErrors,
  FONT_NAMES,
  FONT_SIZES,
  FONT_STYLES,
  FORMAT_TABS,
  HORIZONTALS,
  INDENTED,
  patchesOf,
  pickNumber,
  setHorizontal,
  startOf,
  typeNumber,
  UNDERLINES,
  VERTICALS,
  type FormatDraft,
  type FormatStart,
  type FormatTab,
} from '../format-cells'
import { useMessages } from '../messages'
import type { SheetMessages } from '../contract/messages'
import { useNumberFormats } from '../number-formats'
import FormatBorders from './FormatBorders.vue'
import FormatColour from './FormatColour.vue'
import DialogShell from './DialogShell.vue'
import { fieldValue } from './dom'
import { IconClose } from './icons'
import '../controls.css'

const props = withDefaults(
  defineProps<{
    /** Shown while true; each time it opens, it starts from the active cell's formatting. */
    open: boolean
    /** The payload of the sheet's `state` event: the selection and the active cell's formatting; with null nothing is shown. */
    state: SheetViewState | null
    /** Interface texts; a partial set overrides the English defaults. */
    messages?: Partial<SheetMessages>
    /** The workbook's locale (as for `SheetView`), in which the Number tab names its categories and items; en-US without one. */
    locale?: string
  }>(),
  { messages: undefined, locale: undefined },
)
const emit = defineEmits<{
  /** OK was clicked with changes: patches for `SheetView.format`, a style without a range applying to the selection, and borders on their own ranges. */
  apply: [patches: SheetFormatPatch[]]
  /** OK or Cancel was clicked: the host hides the dialog. */
  close: []
}>()

const m = useMessages(() => props.messages)
const numbers = useNumberFormats(() => props.locale)

const tab = ref<FormatTab>('number')
const start = shallowRef<FormatStart | null>(null)
const draft = ref<FormatDraft | null>(null)
/** Color inputs holding an invalid color, by picker. */
const badColour = reactive<Record<string, boolean>>({})

watch(
  () => props.open && !!props.state,
  (shown) => {
    if (!shown || !props.state) return
    start.value = startOf(props.state)
    draft.value = structuredClone(start.value.draft)
    for (const k of Object.keys(badColour)) delete badColour[k]
  },
  { immediate: true },
)

const errors = computed(() => (draft.value ? draftErrors(draft.value) : {}))
const blocked = computed(() => Object.keys(errors.value).length > 0 || Object.values(badColour).some(Boolean))
/** The category Excel's box names for the code in the input. */
const category = computed(() => (draft.value ? numbers.categoryOf(draft.value.numberCode.trim() || undefined) : null))

function ok(): void {
  if (!start.value || !draft.value || blocked.value) return
  // plain objects: a source in a worker thread clones them, and a reactive proxy cannot be cloned
  const patches = patchesOf(start.value, toRaw(draft.value))
  if (patches.length) emit('apply', patches)
  emit('close')
}
const cancel = () => emit('close')
</script>

<template>
  <DialogShell
    v-if="open && draft && start"
    as="form"
    back-class="xs-fmt-back"
    class="xs-fmt"
    focus="[aria-selected='true']"
    :aria-label="m.formatCells"
    data-fmt
    novalidate
    @submit.prevent="ok"
    @close="cancel"
  >
    <header class="xs-fmt-head">
      <span class="xs-modal-title xs-fmt-title">{{ m.formatCells }}</span>
      <button type="button" class="xs-tool xs-fmt-x" :title="m.cancel" :aria-label="m.cancel" @click="cancel">
        <IconClose />
      </button>
    </header>
    <div class="xs-fmt-tabs" role="tablist">
      <button v-for="t in FORMAT_TABS" :key="t" type="button" class="xs-fmt-tab" role="tab" :aria-selected="tab === t" :data-fmt-tab="t" @click="tab = t">
        {{ m.fmtTab(t) }}
      </button>
    </div>

    <div class="xs-fmt-body" role="tabpanel" :aria-label="m.fmtTab(tab)">
      <div v-show="tab === 'number'" class="xs-fmt-stack">
        <p class="xs-fmt-category" data-fmt-category :data-fmt-category-key="category?.category">
          {{ m.fmtCategory }}: <b>{{ category?.name ?? '' }}</b>
        </p>
        <fieldset class="xs-fmt-box">
          <legend>{{ m.numberFormat }}</legend>
          <div class="xs-fmt-list" role="radiogroup" :aria-label="m.numberFormat">
            <label v-for="c in numbers.choices.value" :key="c.key" class="xs-fmt-choice">
              <input type="radio" name="xs-fmt-number" :checked="draft.numberCode === c.code" :data-fmt-number="c.key" @change="pickNumber(draft, c)" />
              {{ c.name }}
            </label>
          </div>
        </fieldset>
        <label class="xs-fmt-field">
          <span>{{ m.fmtCode }}</span>
          <input
            class="xs-field xs-fmt-input"
            type="text"
            spellcheck="false"
            autocomplete="off"
            :value="draft.numberCode"
            :aria-invalid="!!errors.numberCode"
            data-fmt-code
            @input="typeNumber(draft, fieldValue($event))"
          />
        </label>
        <p v-if="errors.numberCode" class="xs-fmt-error" role="alert">{{ m.fmtCodeEmpty }}</p>
      </div>

      <div v-show="tab === 'alignment'" class="xs-fmt-grid">
        <label class="xs-fmt-field">
          <span>{{ m.fmtHorizontal }}</span>
          <select class="xs-field xs-fmt-select" :value="draft.horizontal" data-fmt-horizontal @change="setHorizontal(draft, fieldValue($event))">
            <option v-for="h in HORIZONTALS" :key="h" :value="h">{{ m.fmtHorizontalKind(h) }}</option>
          </select>
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.fmtIndent }}</span>
          <input v-model="draft.indent" class="xs-field xs-fmt-input" type="text" inputmode="numeric" :disabled="!INDENTED.has(draft.horizontal)" :aria-invalid="!!errors.indent" data-fmt-indent />
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.fmtVertical }}</span>
          <select v-model="draft.vertical" class="xs-field xs-fmt-select" data-fmt-vertical>
            <option v-for="v in VERTICALS" :key="v" :value="v">{{ m.fmtVerticalKind(v) }}</option>
          </select>
        </label>
        <label class="xs-fmt-check">
          <input v-model="draft.wrap" type="checkbox" data-fmt-wrap />
          {{ m.wrapText }}
        </label>
        <p v-if="errors.indent" class="xs-fmt-error" role="alert">{{ m.fmtIndentRange }}</p>
      </div>

      <div v-show="tab === 'font'" class="xs-fmt-grid">
        <label class="xs-fmt-field">
          <span>{{ m.fmtFontName }}</span>
          <input v-model="draft.fontName" class="xs-field xs-fmt-input" type="text" list="xs-fmt-fonts" spellcheck="false" autocomplete="off" :placeholder="m.fmtFontDefault" data-fmt-font-name />
          <datalist id="xs-fmt-fonts">
            <option v-for="f in FONT_NAMES" :key="f" :value="f" />
          </datalist>
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.fmtFontStyle }}</span>
          <select v-model="draft.fontStyle" class="xs-field xs-fmt-select" data-fmt-font-style>
            <option v-for="k in FONT_STYLES" :key="k" :value="k">{{ m.fmtFontStyleKind(k) }}</option>
          </select>
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.fmtSize }}</span>
          <input v-model="draft.fontSize" class="xs-field xs-fmt-input" type="text" inputmode="decimal" list="xs-fmt-sizes" autocomplete="off" placeholder="11" :aria-invalid="!!errors.fontSize" data-fmt-font-size />
          <datalist id="xs-fmt-sizes">
            <option v-for="s in FONT_SIZES" :key="s" :value="String(s)" />
          </datalist>
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.fmtUnderline }}</span>
          <select v-model="draft.underline" class="xs-field xs-fmt-select" data-fmt-underline>
            <option v-for="u in UNDERLINES" :key="u" :value="u">{{ m.fmtUnderlineKind(u) }}</option>
          </select>
        </label>
        <label class="xs-fmt-check">
          <input v-model="draft.strike" type="checkbox" data-fmt-strike />
          {{ m.fmtStrike }}
        </label>
        <p v-if="errors.fontSize" class="xs-fmt-error" role="alert">{{ m.fmtSizeRange }}</p>
        <fieldset class="xs-fmt-box xs-fmt-wide">
          <legend>{{ m.fmtColor }}</legend>
          <FormatColour v-model="draft.fontColour" :none-label="m.colorAuto" :messages="m" name="font" @invalid="badColour.font = $event" />
        </fieldset>
      </div>

      <FormatBorders v-show="tab === 'border'" :start="start" :draft="draft" :messages="m" @invalid="badColour.line = $event" />

      <fieldset v-show="tab === 'fill'" class="xs-fmt-box">
        <legend>{{ m.fmtBackground }}</legend>
        <FormatColour v-model="draft.fill" :none-label="m.fillNone" :messages="m" name="fill" @invalid="badColour.fill = $event" />
      </fieldset>
    </div>

    <footer class="xs-fmt-foot">
      <button type="button" class="xs-button xs-fmt-btn" data-fmt-cancel @click="cancel">{{ m.cancel }}</button>
      <button type="submit" class="xs-button xs-button-main xs-fmt-btn" :disabled="blocked" data-fmt-ok>{{ m.ok }}</button>
    </footer>
  </DialogShell>
</template>

<style src="../format-dialog.css"></style>
