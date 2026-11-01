<script setup lang="ts">
// A timeline as Excel draws it, in a plain neutral look (its style is kept for Excel and not drawn): the header with its
// caption and Clear Filter, the selection's label and the level picker, then the periods of the level under their years
// (or months), the chosen ones filled, scrolled sideways. A click chooses a period, a drag a run of them, Shift+click
// extends the run; the level picker shows years, quarters, months or days.
import { computed, onMounted, ref, watch } from 'vue'
import type { TimelineLevel } from 'xlcalc/contract'
import type { SheetTimeline } from '../contract/source'
import { inSelection, selectionLabel, timelinePeriods } from '../timeline'
import { useSheetContext } from './context'
import { IconDown, IconFilterOff } from './icons'
import { useOutsidePress } from './popup'

const props = defineProps<{ timeline: SheetTimeline }>()
const ctx = useSheetContext()
const { slicers } = ctx
const m = ctx.messages

const LEVELS: TimelineLevel[] = ['years', 'quarters', 'months', 'days']
const level = computed<TimelineLevel>(() => props.timeline.level ?? 'months')
const periods = computed(() => timelinePeriods(props.timeline.bounds, level.value, m.value))
/** The periods' groups above them: a year (a month for days) over the run of its periods. */
const groups = computed(() => {
  const out: { label: string; span: number }[] = []
  for (const p of periods.value) {
    const last = out.at(-1)
    if (last && last.label === p.group) last.span++
    else out.push({ label: p.group, span: 1 })
  }
  return out
})
const label = computed(() => selectionLabel(props.timeline.selection, props.timeline.selectionLevel ?? level.value, m.value))
const usable = computed(() => slicers.canChoose(props.timeline))
const ref_ = () => ({ kind: 'timeline' as const, name: props.timeline.name })

// ── choosing periods: a press starts a run, a drag extends it, the release sends it ──
const run = ref<{ from: number; to: number } | null>(null)
/** The day a Shift+click extends from: the last period pressed, by its first day, so it holds after a change of level. */
let anchor: string | null = null
/** The period of the level that holds a day; -1 for none. */
const periodOf = (day: string) => periods.value.findIndex((p) => p.start.slice(0, 10) <= day.slice(0, 10) && p.end.slice(0, 10) >= day.slice(0, 10))
const chosen = (i: number) => (run.value ? i >= Math.min(run.value.from, run.value.to) && i <= Math.max(run.value.from, run.value.to) : inSelection(periods.value[i]!, props.timeline.selection))
const indexAt = (e: PointerEvent) => Number(document.elementFromPoint(e.clientX, e.clientY)?.closest<HTMLElement>('[data-xs-period]')?.dataset.xsPeriod ?? NaN)
function onDown(e: PointerEvent): void {
  slicers.select(ref_())
  const i = Number((e.target as HTMLElement).closest<HTMLElement>('[data-xs-period]')?.dataset.xsPeriod ?? NaN)
  if (e.button !== 0 || Number.isNaN(i) || !usable.value) return
  ;(e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId)
  // Shift+click extends from the last period pressed, or from the dates chosen before
  const from = e.shiftKey ? periodOf(anchor ?? props.timeline.selection?.start ?? '') : -1
  run.value = { from: from >= 0 ? from : i, to: i }
  if (!e.shiftKey || from < 0) anchor = periods.value[i]!.start
}
function onMove(e: PointerEvent): void {
  if (!run.value) return
  const i = indexAt(e)
  if (!Number.isNaN(i)) run.value = { ...run.value, to: i }
}
function onUp(): void {
  const r = run.value
  if (!r) return
  const a = periods.value[Math.min(r.from, r.to)]
  const b = periods.value[Math.max(r.from, r.to)]
  if (a && b) void slicers.period(props.timeline, a.start, b.end).finally(() => (run.value = null))
  else run.value = null
}

// ── the level picker ──
const picking = ref(false)
const picker = ref<HTMLElement | null>(null)
// a press outside the list closes it, as Esc does
useOutsidePress(picker, () => (picking.value = false), () => picking.value)
function onPickerKey(e: KeyboardEvent): void {
  if (e.key !== 'Escape' || !picking.value) return
  e.preventDefault()
  e.stopPropagation()
  picking.value = false
}
function pickLevel(to: TimelineLevel): void {
  picking.value = false
  void slicers.level(props.timeline, to)
}

/** The periods open at the selection's first, or at the timeline's scroll position. */
const strip = ref<HTMLElement | null>(null)
function scrollToStart(): void {
  const at = props.timeline.selection?.start ?? props.timeline.scrollPosition
  const i = at ? periods.value.findIndex((p) => p.end.slice(0, 10) >= at.slice(0, 10)) : -1
  const el = i >= 0 ? strip.value?.querySelector<HTMLElement>(`[data-xs-period="${i}"]`) : null
  if (el && strip.value) strip.value.scrollLeft = el.offsetLeft
}
onMounted(scrollToStart)
watch(level, () => requestAnimationFrame(scrollToStart))
</script>

<template>
  <div class="xs-timeline" :data-xs-timeline="timeline.name">
    <div v-if="timeline.showHeader !== false" class="xs-slicer-head">
      <span class="xs-slicer-caption">{{ timeline.caption ?? timeline.name }}</span>
      <button type="button" class="xs-slicer-tool" :title="m.slicerClearFilter" :aria-label="m.slicerClearFilter" :disabled="!timeline.selection || !usable" data-xs-timeline-clear @click="slicers.clear(ref_())">
        <IconFilterOff />
      </button>
    </div>
    <div class="xs-timeline-bar">
      <span v-if="timeline.showSelectionLabel !== false" class="xs-timeline-label" data-xs-timeline-label>{{ label }}</span>
      <span v-if="timeline.showLevel !== false" ref="picker" class="xs-timeline-level" @keydown="onPickerKey">
        <button type="button" class="xs-timeline-level-button" :aria-expanded="picking" :disabled="!slicers.canChoose(timeline)" data-xs-timeline-level @click="picking = !picking">
          {{ m.timelineLevel(level) }}<IconDown />
        </button>
        <span v-if="picking" class="xs-timeline-levels" role="listbox">
          <button v-for="l in LEVELS" :key="l" type="button" role="option" :aria-selected="l === level" :class="['xs-timeline-level-item', { 'xs-on': l === level }]" :data-xs-level="l" @click="pickLevel(l)">
            {{ m.timelineLevel(l) }}
          </button>
        </span>
      </span>
    </div>
    <div ref="strip" :class="['xs-timeline-strip', { 'xs-timeline-noscroll': timeline.showScrollbar === false }]">
      <div class="xs-timeline-grid" :style="{ gridTemplateColumns: `repeat(${periods.length}, var(--xs-timeline-tile, 40px))` }" data-xs-periods @pointerdown="onDown" @pointermove="onMove" @pointerup="onUp" @pointercancel="run = null">
        <span v-for="(g, i) in level === 'years' ? [] : groups" :key="`g${i}`" class="xs-timeline-group" :style="{ gridColumn: `span ${g.span}` }">{{ g.label }}</span>
        <span v-for="(p, i) in periods" :key="p.start" :class="['xs-timeline-period', { 'xs-on': chosen(i) }]" :title="p.label" :data-xs-period="i" :aria-selected="chosen(i)">
          <span class="xs-timeline-tile">{{ p.label }}</span>
          <span class="xs-timeline-fill"></span>
        </span>
      </div>
    </div>
  </div>
</template>
