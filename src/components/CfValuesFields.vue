<script setup lang="ts">
// "Format all cells based on their values" in Excel's rule dialog: a two- or three-color scale's points, a data bar's bounds
// and appearance, an icon set's style and thresholds, each threshold a type and a value (a number, or a formula after `=`).
import { computed } from 'vue'
import type { SheetMessages } from '../contract/messages'
import {
  BAR_HIGH_TYPES,
  BAR_LOW_TYPES,
  HIGH_TYPES,
  ICON_SETS,
  ICON_TYPES,
  LOW_TYPES,
  MID_TYPES,
  pickIconSet,
  VALUES_STYLES,
  type PointDraft,
  type RuleDraft,
  type ValueType,
} from '../core/cf-draft'
import { iconCount } from '../core/cf-presets'
import { CfIcon } from './cf-icons'
import { fieldValue } from './dom'

const props = defineProps<{ draft: RuleDraft; messages: SheetMessages }>()
const m = computed(() => props.messages)
const d = computed(() => props.draft)

/** Types whose box takes no value: the lowest, the highest, automatic. */
const NO_VALUE: ReadonlySet<ValueType> = new Set(['min', 'max', 'autoMin', 'autoMax'])
/** A scale's columns: their caption and the types each offers. */
const columns = computed(() => {
  const pts = d.value.values === 'scale2' ? d.value.scale2 : d.value.scale3
  const which = pts.length === 3 ? (['min', 'mid', 'max'] as const) : (['min', 'max'] as const)
  return which.map((w, i) => ({ w, p: pts[i]!, types: w === 'min' ? LOW_TYPES : w === 'mid' ? MID_TYPES : HIGH_TYPES }))
})
/** A new type: the value box empties for one without a value. */
function setType(p: PointDraft | { type: ValueType; value: string }, t: string): void {
  p.type = t as ValueType
  if (NO_VALUE.has(p.type)) p.value = ''
}
const barPoints = computed(() => [
  { w: 'min' as const, p: d.value.barMin, types: BAR_LOW_TYPES },
  { w: 'max' as const, p: d.value.barMax, types: BAR_HIGH_TYPES },
])
const DIRECTIONS = ['context', 'leftToRight', 'rightToLeft'] as const
const upper = (e: Event) => fieldValue(e).toUpperCase()
/** The icon rows as Excel lists them, the highest first; each threshold's icon, reversed with Reverse Icon Order. */
const iconRows = computed(() => {
  const n = iconCount(d.value.iconSet)
  const icon = (i: number) => ({ set: d.value.iconSet, index: d.value.iconReverse ? n - 1 - i : i })
  return d.value.iconPoints.map((p, k) => ({ p, k, icon: icon(k + 1) })).reverse()
})
const lowIcon = computed(() => ({ set: d.value.iconSet, index: d.value.iconReverse ? iconCount(d.value.iconSet) - 1 : 0 }))
</script>

<template>
  <div class="xs-cf-values">
    <label class="xs-cf-line">
      <span>{{ m.cfFormatStyle }}</span>
      <select v-model="draft.values" class="xs-field" data-cf-values>
        <option v-for="v in VALUES_STYLES" :key="v" :value="v">{{ m.cfStyleKind(v) }}</option>
      </select>
    </label>

    <div v-if="draft.values === 'scale2' || draft.values === 'scale3'" class="xs-cf-points" :data-cf-points="columns.length">
      <div v-for="c in columns" :key="c.w" class="xs-cf-point">
        <b>{{ m.cfPoint(c.w) }}</b>
        <label class="xs-fmt-field">
          <span>{{ m.cfTypeLabel }}</span>
          <select class="xs-field" :value="c.p.type" :data-cf-type="c.w" @change="setType(c.p, fieldValue($event))">
            <option v-for="t in c.types" :key="t" :value="t">{{ m.cfValueType(t) }}</option>
          </select>
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.cfValueLabel }}</span>
          <input v-model="c.p.value" class="xs-field" type="text" spellcheck="false" autocomplete="off" :disabled="NO_VALUE.has(c.p.type)" :data-cf-point-value="c.w" />
        </label>
        <label class="xs-fmt-field">
          <span>{{ m.cfColorLabel }}</span>
          <input class="xs-cf-colour" type="color" :value="c.p.color" :data-cf-color="c.w" @input="c.p.color = upper($event)" />
        </label>
      </div>
    </div>

    <template v-else-if="draft.values === 'dataBar'">
      <div class="xs-cf-points" data-cf-points="2">
        <div v-for="c in barPoints" :key="c.w" class="xs-cf-point">
          <b>{{ m.cfPoint(c.w) }}</b>
          <label class="xs-fmt-field">
            <span>{{ m.cfTypeLabel }}</span>
            <select class="xs-field" :value="c.p.type" :data-cf-type="c.w" @change="setType(c.p, fieldValue($event))">
              <option v-for="t in c.types" :key="t" :value="t">{{ m.cfValueType(t) }}</option>
            </select>
          </label>
          <label class="xs-fmt-field">
            <span>{{ m.cfValueLabel }}</span>
            <input v-model="c.p.value" class="xs-field" type="text" spellcheck="false" autocomplete="off" :disabled="NO_VALUE.has(c.p.type)" :data-cf-point-value="c.w" />
          </label>
        </div>
      </div>
      <p class="xs-cf-caption">{{ m.cfBarAppearance }}</p>
      <div class="xs-cf-grid">
        <select v-model="draft.barGradient" class="xs-field" data-cf-bar-fill>
          <option :value="true">{{ m.cfBarFill(true) }}</option>
          <option :value="false">{{ m.cfBarFill(false) }}</option>
        </select>
        <input class="xs-cf-colour" type="color" :value="draft.barColor" :aria-label="m.cfColorLabel" data-cf-bar-color @input="draft.barColor = upper($event)" />
        <select v-model="draft.barBorder" class="xs-field" data-cf-bar-border>
          <option :value="true">{{ m.cfBarBorder(true) }}</option>
          <option :value="false">{{ m.cfBarBorder(false) }}</option>
        </select>
        <input class="xs-cf-colour" type="color" :value="draft.barBorderColor" :disabled="!draft.barBorder" :aria-label="m.cfColorLabel" data-cf-bar-border-color @input="draft.barBorderColor = upper($event)" />
      </div>
      <label class="xs-cf-line">
        <span>{{ m.cfDirection }}</span>
        <select v-model="draft.barDirection" class="xs-field" data-cf-bar-direction>
          <option v-for="k in DIRECTIONS" :key="k" :value="k">{{ m.cfDirectionKind(k) }}</option>
        </select>
      </label>
      <label class="xs-fmt-check">
        <input v-model="draft.barOnly" type="checkbox" data-cf-bar-only />
        {{ m.cfShowBarOnly }}
      </label>
    </template>

    <template v-else>
      <label class="xs-cf-line">
        <span>{{ m.cfIconStyle }}</span>
        <select class="xs-field" :value="draft.iconSet" data-cf-icon-set @change="pickIconSet(draft, fieldValue($event) as RuleDraft['iconSet'])">
          <option v-for="s in ICON_SETS" :key="s" :value="s">{{ m.cfIconSet(s) }}</option>
        </select>
      </label>
      <div class="xs-cf-grid">
        <label class="xs-fmt-check">
          <input v-model="draft.iconReverse" type="checkbox" data-cf-icon-reverse />
          {{ m.cfReverse }}
        </label>
        <label class="xs-fmt-check">
          <input v-model="draft.iconOnly" type="checkbox" data-cf-icon-only />
          {{ m.cfShowIconOnly }}
        </label>
      </div>
      <p class="xs-cf-caption">{{ m.cfIconRules }}</p>
      <div v-for="r in iconRows" :key="r.k" class="xs-cf-icon-row" :data-cf-icon-row="r.k">
        <CfIcon :icon="r.icon" />
        <span>{{ m.cfWhenValue }}</span>
        <select v-model="r.p.gte" class="xs-field" :data-cf-gte="r.k">
          <option :value="true">&gt;=</option>
          <option :value="false">&gt;</option>
        </select>
        <input v-model="r.p.value" class="xs-field xs-cf-icon-value" type="text" spellcheck="false" autocomplete="off" :data-cf-icon-value="r.k" />
        <select class="xs-field" :value="r.p.type" :data-cf-icon-type="r.k" @change="setType(r.p, fieldValue($event))">
          <option v-for="t in ICON_TYPES" :key="t" :value="t">{{ m.cfValueType(t) }}</option>
        </select>
      </div>
      <div class="xs-cf-icon-row">
        <CfIcon :icon="lowIcon" />
        <span>{{ m.cfIconLast }}</span>
      </div>
    </template>
  </div>
</template>
