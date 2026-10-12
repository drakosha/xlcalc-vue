<script setup lang="ts">
// Excel's alert for a value a rule of data validation does not take: its style's icon, the rule's title ("Microsoft Excel"
// without one) and message, and the style's buttons: Stop, Retry and Cancel; Warning, "Continue?" and Yes, No and Cancel;
// Information, OK and Cancel. It opens on Excel's default button (Retry, No, OK), Esc is Cancel; 44 px targets on a phone.
import { computed, nextTick, ref, useId, watch } from 'vue'
import { ALERT_BUTTONS, ALERT_DEFAULT, type AlertAnswer, type AlertButton } from '../core/validation'
import { enMessages, type SheetMessages } from '../messages'
import type { SheetValidationAlert } from '../source'

const props = withDefaults(
  defineProps<{
    /** The alert to show; null shows nothing. */
    alert: SheetValidationAlert | null
    messages?: Partial<SheetMessages>
  }>(),
  { messages: undefined },
)
const emit = defineEmits<{
  /** The person's answer: `accept` (Yes, OK), `retry` (Retry, No) or `cancel`. */
  answer: [answer: AlertAnswer]
}>()

const m = computed<SheetMessages>(() => ({ ...enMessages, ...props.messages }))
const box = ref<HTMLElement | null>(null)
const id = useId()
const buttons = computed(() => (props.alert ? ALERT_BUTTONS[props.alert.severity] : []))
const label = (b: AlertButton): string =>
  ({ retry: m.value.validationRetry, cancel: m.value.cancel, yes: m.value.validationYes, no: m.value.validationNo, ok: m.value.ok })[b]

// Each alert opens on its default button, as Excel's.
watch(
  () => props.alert,
  async (a) => {
    if (!a) return
    await nextTick()
    box.value?.querySelector<HTMLElement>(`[data-xs-alert-button="${ALERT_DEFAULT[a.severity]}"]`)?.focus()
  },
  { immediate: true },
)
/** Tab goes round the buttons, the alert being modal. */
function onKey(e: KeyboardEvent): void {
  if (e.key === 'Escape') {
    e.preventDefault()
    emit('answer', 'cancel')
    return
  }
  if (e.key !== 'Tab') return
  const all = [...(box.value?.querySelectorAll<HTMLElement>('button') ?? [])]
  const at = all.indexOf(document.activeElement as HTMLElement)
  e.preventDefault()
  all[(at + (e.shiftKey ? all.length - 1 : 1)) % all.length]?.focus()
}
</script>

<template>
  <div v-if="alert" class="xs-va-back" data-xs-alert-back>
    <div
      ref="box"
      class="xs-va"
      role="alertdialog"
      aria-modal="true"
      :aria-labelledby="`${id}-title`"
      :aria-describedby="`${id}-text`"
      :data-xs-alert="alert.severity"
      @keydown.stop="onKey"
    >
      <header :id="`${id}-title`" class="xs-va-title">{{ alert.title || m.validationTitle }}</header>
      <div class="xs-va-body">
        <svg class="xs-va-icon" viewBox="0 0 32 32" role="img" :aria-label="m.validationStyle(alert.severity)">
          <template v-if="alert.severity === 'stop'">
            <circle cx="16" cy="16" r="15" fill="var(--xs-alert-stop)" />
            <path d="M11 11l10 10M21 11 11 21" stroke="#fff" stroke-width="2.6" stroke-linecap="round" />
          </template>
          <template v-else-if="alert.severity === 'warning'">
            <path d="M16 2.5 30.5 28.5h-29z" fill="var(--xs-alert-warning)" stroke-linejoin="round" />
            <path d="M16 11v9" stroke="#1f2328" stroke-width="2.8" stroke-linecap="round" />
            <circle cx="16" cy="24.2" r="1.7" fill="#1f2328" />
          </template>
          <template v-else>
            <circle cx="16" cy="16" r="15" fill="var(--xs-alert-information)" />
            <circle cx="16" cy="9.5" r="1.9" fill="#fff" />
            <path d="M16 14v10" stroke="#fff" stroke-width="2.8" stroke-linecap="round" />
          </template>
        </svg>
        <div :id="`${id}-text`" class="xs-va-text">
          <p data-xs-alert-message>{{ alert.message }}</p>
          <p v-if="alert.severity === 'warning'" data-xs-alert-question>{{ m.validationContinue }}</p>
        </div>
      </div>
      <footer class="xs-va-foot">
        <button
          v-for="b in buttons"
          :key="b.button"
          type="button"
          :class="['xs-va-btn', { 'xs-va-main': b.button === ALERT_DEFAULT[alert.severity] }]"
          :data-xs-alert-button="b.button"
          @click="emit('answer', b.answer)"
        >
          {{ label(b.button) }}
        </button>
      </footer>
    </div>
  </div>
</template>

<style src="../validation.css"></style>
