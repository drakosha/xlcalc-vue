<script setup lang="ts">
// Excel's Find and Replace below the formula bar: the query, Find Next and Find Previous, the dialog's options, the range it
// searches within, and the Replace row; the source finds, matches in the viewport are highlighted and the current one pulses.
import type { FindRunner } from '../core/find'
import { useSheetContext } from './context'
import { fieldValue, refInto } from './dom'
import { IconClose, IconDown, IconOptions, IconUp } from './icons'
import '../controls.css'

const ctx = useSheetContext()
const { find } = ctx
const m = ctx.messages
const setInput = refInto(find.input)
/** Replacements go to the source as an edit does: their cells drawn from the outcome, a rejection in the sheet's message line. */
const run: FindRunner = (call) => ctx.ed.run(call, [], null, null)
</script>

<template>
  <div class="xs-find" role="search">
    <div class="xs-find-row">
      <button
        v-if="find.canReplace.value"
        type="button"
        :class="['xs-tool', 'xs-find-toggle', { 'xs-tool-on': find.replacing.value }]"
        :title="m.replaceToggle"
        :aria-label="m.replaceToggle"
        :aria-expanded="find.replacing.value"
        data-xs-find-replace
        @click="find.replacing.value = !find.replacing.value"
      >
        <IconDown />
      </button>
      <input
        :ref="setInput"
        v-model="find.text.value"
        class="xs-field xs-field-m xs-find-input"
        type="search"
        spellcheck="false"
        :placeholder="m.findPlaceholder"
        @keydown="find.onFindKey"
      />
      <span class="xs-find-status" aria-live="polite">{{ find.status.value }}</span>
      <span class="xs-find-nav">
        <button type="button" class="xs-tool" :title="m.findPrev" :aria-label="m.findPrev" :disabled="!find.hits.value.length" @click="find.step(-1)">
          <IconUp />
        </button>
        <button type="button" class="xs-tool" :title="m.findNext" :aria-label="m.findNext" :disabled="!find.hits.value.length" @click="find.step(1)">
          <IconDown />
        </button>
        <button
          type="button"
          :class="['xs-tool', { 'xs-tool-on': find.optionsOpen.value }]"
          :title="m.findOptions"
          :aria-label="m.findOptions"
          :aria-expanded="find.optionsOpen.value"
          data-xs-find-options
          @click="find.optionsOpen.value = !find.optionsOpen.value"
        >
          <IconOptions />
        </button>
        <button type="button" class="xs-tool" :title="m.close" :aria-label="m.close" @click="find.closeFind()">
          <IconClose />
        </button>
      </span>
    </div>
    <div v-if="find.replacing.value && find.canReplace.value" class="xs-find-row">
      <input v-model="find.by.value" class="xs-field xs-field-m xs-find-input" type="text" spellcheck="false" :placeholder="m.replacePlaceholder" :aria-label="m.replacePlaceholder" data-xs-replace-input @keydown="find.onFindKey" />
      <button type="button" class="xs-button" :disabled="!find.text.value || find.writing.value" data-xs-replace-one @click="find.replaceOne(run)">{{ m.replaceOne }}</button>
      <button type="button" class="xs-button" :disabled="!find.text.value || find.writing.value" data-xs-replace-all @click="find.replaceAll(run)">{{ m.replaceAll }}</button>
    </div>
    <div v-if="find.optionsOpen.value || find.scope.value" class="xs-find-row xs-find-more">
      <span v-if="find.scope.value" class="xs-find-scope" data-xs-find-scope>
        {{ m.findWithin(find.scope.value) }}
        <button type="button" class="xs-tool" :title="m.findWholeSheet" :aria-label="m.findWholeSheet" @click="find.scope.value = null">
          <IconClose />
        </button>
      </span>
      <template v-if="find.optionsOpen.value">
        <label class="xs-find-opt">
          <input v-model="find.options.matchCase" type="checkbox" data-xs-find-case />
          {{ m.findMatchCase }}
        </label>
        <label class="xs-find-opt">
          <input v-model="find.options.wholeCell" type="checkbox" data-xs-find-whole />
          {{ m.findWholeCell }}
        </label>
        <label class="xs-find-opt">
          {{ m.findLookIn }}
          <select class="xs-field xs-field-s" :value="find.lookIn.value" :disabled="find.replacing.value" data-xs-find-look @change="find.options.lookIn = fieldValue($event) === 'values' ? 'values' : 'formulas'">
            <option value="formulas">{{ m.findInFormulas }}</option>
            <option value="values">{{ m.findInValues }}</option>
          </select>
        </label>
        <label class="xs-find-opt">
          {{ m.findOrder }}
          <select class="xs-field xs-field-s" :value="find.options.byColumns ? 'columns' : 'rows'" data-xs-find-order @change="find.options.byColumns = fieldValue($event) === 'columns'">
            <option value="rows">{{ m.findByRows }}</option>
            <option value="columns">{{ m.findByColumns }}</option>
          </select>
        </label>
      </template>
    </div>
  </div>
</template>
