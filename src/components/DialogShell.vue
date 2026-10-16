<script setup lang="ts">
// A modal box over the page, as the sheet's dialogs and alerts are: the dimmed backdrop (a press on it closes the box
// where `outside` says so), Esc to close, the focus put inside on opening and kept inside by Tab. The box's own attributes
// and listeners (its class, label, a form's submit) are those given to the shell.
import { nextTick, onMounted, ref } from 'vue'

defineOptions({ inheritAttrs: false })
const props = withDefaults(
  defineProps<{
    /** The box's element: a form where Enter submits. */
    as?: 'div' | 'form'
    role?: 'dialog' | 'alertdialog'
    /** A press on the backdrop closes the box; Excel's alerts stay up. */
    outside?: boolean
    /** Full screen on a phone (the sheet's .xs-narrow), as Excel's larger windows go. */
    full?: boolean
    /** A class of the backdrop's own: the Format Cells dialog's carries the sheet's variables outside a sheet. */
    backClass?: string
    /** Where the focus goes on opening: a selector inside the box (a field gets its text selected); the box itself without one. */
    focus?: string
  }>(),
  { as: 'div', role: 'dialog', outside: true, full: false, backClass: undefined, focus: undefined },
)
const emit = defineEmits<{
  /** Esc, or a press on the backdrop where `outside`. */
  close: []
}>()

const box = ref<HTMLElement | null>(null)
const FOCUSABLE = 'button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [href], [tabindex]:not([tabindex="-1"])'

onMounted(async () => {
  await nextTick()
  const el = (props.focus && box.value?.querySelector<HTMLElement>(props.focus)) || box.value
  el?.focus()
  if (el instanceof HTMLInputElement) el.select()
})

/** Esc closes; Tab past the last control goes round to the first and Shift+Tab back, the box being modal. */
function onKey(e: KeyboardEvent): void {
  if (e.key === 'Escape') {
    e.preventDefault()
    e.stopPropagation()
    emit('close')
    return
  }
  if (e.key !== 'Tab' || !box.value) return
  const all = [...box.value.querySelectorAll<HTMLElement>(FOCUSABLE)]
  const at = all.indexOf(document.activeElement as HTMLElement)
  const to = e.shiftKey ? (at <= 0 ? all.length - 1 : -1) : at < 0 || at === all.length - 1 ? 0 : -1
  const next = to < 0 ? undefined : all[to]
  if (!next) return
  e.preventDefault()
  next.focus()
}
</script>

<template>
  <div :class="['xs-modal-back', backClass, { 'xs-modal-full': full, 'xs-modal-alert': role === 'alertdialog' }]" @pointerdown.self="outside && emit('close')">
    <component :is="as" ref="box" v-bind="$attrs" class="xs-modal" :role="role" aria-modal="true" tabindex="-1" @keydown="onKey">
      <slot />
    </component>
  </div>
</template>

<style src="../modal.css"></style>
