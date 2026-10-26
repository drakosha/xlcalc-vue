<script setup lang="ts">
// A conditional format's preview, as Excel's Rules Manager and rule dialogs show it: sample text in the rule's format,
// a color scale's gradient, a data bar, an icon set's icons; "No Format Set" for a rule without one.
import { computed } from 'vue'
import type { ConditionalFormat, ConditionalStyle } from 'xlcalc/contract'
import type { SheetMessages } from '../contract/messages'
import { borderCss, setSide } from '../core/cell-style'
import { iconCount } from '../core/cf-presets'
import { CfIcon } from './cf-icons'

const props = defineProps<{ rule: ConditionalFormat | null; messages: SheetMessages }>()

/** The sample's CSS for a differential format: its font, fill and borders. */
function styleCss(st: ConditionalStyle): Record<string, string> {
  const css: Record<string, string> = {}
  const f = st.font
  if (f?.bold) css.fontWeight = '700'
  if (f?.italic) css.fontStyle = 'italic'
  const lines = [f?.underline ? 'underline' : '', f?.strike ? 'line-through' : ''].filter(Boolean)
  if (lines.length) css.textDecoration = lines.join(' ')
  if (f?.underline === 'double') css.textDecorationStyle = 'double'
  if (f?.color) css.color = f.color
  if (st.fill) css.background = st.fill.color
  for (const [side, key] of [['Top', 'top'], ['Right', 'right'], ['Bottom', 'bottom'], ['Left', 'left']] as const) {
    const line = borderCss(st.border?.[key])
    if (line) setSide(css, side, line)
  }
  return css
}

const look = computed(() => {
  const r = props.rule
  if (!r) return { kind: 'none' as const }
  if (r.type === 'colorScale') return { kind: 'scale' as const, css: { background: `linear-gradient(to right, ${r.points.map((p) => p.color).join(', ')})` } }
  if (r.type === 'dataBar') {
    const fill = r.gradient === false ? r.color : `linear-gradient(to right, ${r.color}, color-mix(in srgb, ${r.color} 18%, white))`
    return { kind: 'bar' as const, css: { background: fill, ...(r.border ? { border: `1px solid ${r.border}` } : {}) } }
  }
  if (r.type === 'iconSet') {
    const n = iconCount(r.icons)
    const order = Array.from({ length: n }, (_, i) => (r.reverse ? n - 1 - i : i))
    return { kind: 'icons' as const, icons: order.map((index) => ({ set: r.icons, index })) }
  }
  const st = r.style ?? {}
  if (!st.font && !st.fill && !st.border && st.numberFormat === undefined) return { kind: 'none' as const }
  return { kind: 'style' as const, css: styleCss(st) }
})
</script>

<template>
  <span class="xs-cf-preview" :data-cf-preview="look.kind">
    <span v-if="look.kind === 'none'" class="xs-cf-sample xs-muted">{{ messages.cfNoFormat }}</span>
    <span v-else-if="look.kind === 'style'" class="xs-cf-sample" :style="look.css">{{ messages.cfSample }}</span>
    <span v-else-if="look.kind === 'icons'" class="xs-cf-sample xs-cf-icons">
      <CfIcon v-for="icon in look.icons" :key="icon.index" :icon="icon" />
    </span>
    <span v-else class="xs-cf-sample xs-cf-band"><span :class="look.kind === 'bar' ? 'xs-cf-bar' : 'xs-cf-scale'" :style="look.css" /></span>
  </span>
</template>
