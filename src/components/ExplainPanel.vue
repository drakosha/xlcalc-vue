<script setup lang="ts">
// Excel's Trace Precedents as a panel right of the grid (below it on a phone): the source's tree of the active cell's
// precedents, its references going to their ranges, where the panel follows; the cross and Esc close it.
import type { ExplainPart } from '../core/explain'
import { useSheetContext } from './context'
import { IconClose, IconSpinner } from './icons'
import '../controls.css'

const ctx = useSheetContext()
const { explain } = ctx
const m = ctx.messages

/** A reference to a sheet the book has goes there; any other stays text. */
const goes = (part: ExplainPart) => !!part.ref && ctx.f.refSheet(part.ref) >= 0
const title = (part: ExplainPart) => (part.ref && part.ref.sheet !== ctx.sheetName.value ? m.value.goSheet(part.ref.sheet) : m.value.showRange)
function go(part: ExplainPart): void {
  if (part.ref) void ctx.nav.goRef(part.ref)
}
function onKey(e: KeyboardEvent): void {
  if (e.key !== 'Escape') return
  e.preventDefault()
  explain.close()
}
</script>

<template>
  <aside class="xs-explain" :aria-label="m.explainTitle" data-xs-explain @keydown="onKey">
    <div class="xs-explain-head">
      <span class="xs-explain-title">{{ m.explainTitle }}</span>
      <span class="xs-explain-place" :title="explain.place.value" data-xs-explain-place>{{ explain.place.value }}</span>
      <IconSpinner v-if="explain.busy.value" class="xs-spin xs-explain-busy" />
      <button type="button" class="xs-tool" :title="m.close" :aria-label="m.close" data-xs-explain-close @click="explain.close()">
        <IconClose />
      </button>
    </div>
    <div class="xs-explain-body" data-xs-explain-body>
      <p v-if="explain.failed.value" class="xs-explain-note" role="alert">{{ m.explainFailed }}</p>
      <p v-else-if="!explain.lines.value.length" class="xs-explain-note">{{ explain.busy.value ? m.explainReading : m.explainEmpty }}</p>
      <div v-for="(line, i) in explain.lines.value" :key="i" class="xs-explain-line" :style="{ '--xs-explain-depth': line.depth }" data-xs-explain-line>
        <template v-for="(part, k) in line.parts" :key="k">
          <button v-if="goes(part)" type="button" class="xs-explain-ref" :title="title(part)" :data-xs-explain-ref="part.text" @click="go(part)">{{ part.text }}</button>
          <span v-else :class="part.from ? 'xs-explain-from' : undefined">{{ part.text }}</span>
        </template>
      </div>
    </div>
  </aside>
</template>

<style src="../explain.css"></style>
