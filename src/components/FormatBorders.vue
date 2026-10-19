<script setup lang="ts">
// The Border tab of the Format Cells dialog, as Excel's: a line's kind and colour, the presets None, Outline and Inside,
// a button for each edge, and a preview of the selection's edges drawn as the grid draws them.
import { computed, h, type FunctionalComponent } from 'vue'
import type { BorderLine } from 'xlcalc/contract'
import { borderCss, setSide } from '../core/cell-style'
import { edgesIn, LINE_STYLES, pressEdge, pressPreset, shownEdge, type FormatDraft, type FormatStart } from '../format-cells'
import type { SheetMessages } from '../messages'
import { upperColour, type BorderEdge } from '../toolbar'
import FormatColour from './FormatColour.vue'

const props = defineProps<{ start: FormatStart; draft: FormatDraft; messages: SheetMessages }>()
const emit = defineEmits<{ invalid: [bad: boolean] }>()
const m = computed(() => props.messages)

const PRESETS = ['none', 'outline', 'inside'] as const
/** The edge buttons as Excel lays them round the preview: across on its left, down under it. */
const EDGE_GROUPS: { cls: string; edges: BorderEdge[] }[] = [
  { cls: 'xs-fmt-edge-col', edges: ['top', 'insideH', 'bottom'] },
  { cls: 'xs-fmt-edge-row', edges: ['left', 'insideV', 'right'] },
]
const EDGE_PATH: Record<BorderEdge, string> = { top: 'M4 4h16', insideH: 'M4 12h16', bottom: 'M4 20h16', left: 'M4 4v16', insideV: 'M12 4v16', right: 'M20 4v16' }
const PRESET_PATHS: Record<(typeof PRESETS)[number], string[]> = { none: [], outline: ['M4 4h16v16H4z'], inside: ['M4 12h16', 'M12 4v16'] }
/** A preset's or an edge's icon: the cell's frame dashed and the lines it draws. */
const BorderIcon: FunctionalComponent<{ paths: string[] }> = (p) =>
  h('svg', { class: 'xs-icon', viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 2, 'aria-hidden': 'true' }, [
    h('path', { d: 'M4 4h16v16H4z', 'stroke-width': 1, 'stroke-dasharray': '2 2' }),
    ...p.paths.map((d) => h('path', { d })),
  ])
BorderIcon.props = ['paths']

const has = computed(() => new Set(edgesIn(props.start)))
const DASHES: Partial<Record<BorderLine['style'], string>> = { hair: '1 1', dotted: '1 2', dashed: '4 2', mediumDashed: '6 3', dashDot: '5 2 1 2', mediumDashDot: '7 3 2 3', dashDotDot: '5 2 1 2 1 2', mediumDashDotDot: '7 3 2 3 2 3', slantDashDot: '7 2 3 2' }
/** Each line's sample as an SVG stroke: its width and dashes, a double one as two strokes. */
const SAMPLES = Object.fromEntries(
  LINE_STYLES.map((style) => [style, { width: style === 'double' ? 1 : Math.min(3, parseInt(borderCss({ style })!.width, 10)), dash: DASHES[style] ?? '', double: style === 'double' }]),
) as Record<BorderLine['style'], { width: number; dash: string; double: boolean }>
function setColour(c: string | null): void {
  props.draft.line = c ? { style: props.draft.line.style, color: c } : { style: props.draft.line.style }
}

/** The preview's cells: two by two for a range (its outline and the lines inside), one for a cell. */
const multi = computed(() => has.value.has('insideH') || has.value.has('insideV'))
const cells = computed(() => {
  const rows = has.value.has('insideH') ? 2 : 1
  const cols = has.value.has('insideV') ? 2 : 1
  const out: { key: string; style: Record<string, string> }[] = []
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const css: Record<string, string> = {}
      const side = (s: 'Top' | 'Right' | 'Bottom' | 'Left', e: BorderEdge) => {
        const line = shownEdge(props.start, props.draft, e)
        if (line === undefined) setSide(css, s, { width: '1px', style: 'dotted', color: 'var(--xs-grid)' })
        else if (line) setSide(css, s, borderCss(line)!)
      }
      side('Top', r === 0 ? 'top' : 'insideH')
      side('Left', c === 0 ? 'left' : 'insideV')
      if (r === rows - 1) side('Bottom', 'bottom')
      if (c === cols - 1) side('Right', 'right')
      out.push({ key: `${r}:${c}`, style: css })
    }
  }
  return out
})
const pressed = (e: BorderEdge) => !!shownEdge(props.start, props.draft, e)
</script>

<template>
  <div class="xs-fmt-grid">
    <fieldset class="xs-fmt-box">
      <legend>{{ m.fmtLine }}</legend>
      <div class="xs-fmt-lines" role="radiogroup" :aria-label="m.fmtLine">
        <button
          v-for="st in LINE_STYLES"
          :key="st"
          type="button"
          class="xs-fmt-line"
          role="radio"
          :aria-checked="draft.line.style === st"
          :aria-label="m.fmtLineStyle(st)"
          :title="m.fmtLineStyle(st)"
          :data-fmt-line="st"
          @click="draft.line = { ...draft.line, style: st }"
        >
          <svg viewBox="0 0 60 8" preserveAspectRatio="none" aria-hidden="true">
            <template v-if="SAMPLES[st].double">
              <line x1="2" y1="2.5" x2="58" y2="2.5" stroke="currentColor" stroke-width="1" />
              <line x1="2" y1="5.5" x2="58" y2="5.5" stroke="currentColor" stroke-width="1" />
            </template>
            <line v-else x1="2" y1="4" x2="58" y2="4" stroke="currentColor" :stroke-width="SAMPLES[st].width" :stroke-dasharray="SAMPLES[st].dash" />
          </svg>
        </button>
      </div>
      <span class="xs-fmt-label">{{ m.fmtColor }}</span>
      <FormatColour :model-value="upperColour(draft.line.color)" :none-label="m.colorAuto" :messages="m" name="line" @update:model-value="setColour" @invalid="emit('invalid', $event)" />
    </fieldset>
    <div class="xs-fmt-stack">
      <fieldset class="xs-fmt-box">
        <legend>{{ m.fmtPresets }}</legend>
        <div class="xs-fmt-presets">
          <button v-for="p in PRESETS" :key="p" type="button" class="xs-fmt-preset" :disabled="p === 'inside' && !multi" :data-fmt-preset="p" @click="pressPreset(start, draft, p)">
            <BorderIcon :paths="PRESET_PATHS[p]" />
            {{ m.fmtPreset(p) }}
          </button>
        </div>
      </fieldset>
      <fieldset class="xs-fmt-box">
        <legend>{{ m.fmtEdges }}</legend>
        <div class="xs-fmt-edges">
          <div v-for="g in EDGE_GROUPS" :key="g.cls" :class="g.cls">
            <button v-for="e in g.edges" :key="e" type="button" class="xs-fmt-edge" :disabled="!has.has(e)" :aria-pressed="pressed(e)" :aria-label="m.fmtEdge(e)" :title="m.fmtEdge(e)" :data-fmt-edge="e" @click="pressEdge(start, draft, e)">
              <BorderIcon :paths="[EDGE_PATH[e]]" />
            </button>
          </div>
          <div class="xs-fmt-preview" :style="{ gridTemplateColumns: has.has('insideV') ? '1fr 1fr' : '1fr' }" role="img" :aria-label="m.fmtEdges" data-fmt-preview>
            <span v-for="c in cells" :key="c.key" class="xs-fmt-pcell" :style="c.style">Aa</span>
          </div>
        </div>
      </fieldset>
    </div>
  </div>
</template>
