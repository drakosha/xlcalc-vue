<script setup lang="ts">
// Excel's alert for a value a rule of data validation does not take: its style's icon, the rule's title ("Microsoft Excel"
// without one) and message, and the style's buttons: Stop, Retry and Cancel; Warning, "Continue?" and Yes, No and Cancel;
// Information, OK and Cancel; the sheet's question on a formula that does not read, Fix and Restore; Excel's correction of a
// formula left open, Yes and No, and its message after No, OK; the warning before a merge, OK and Cancel. It opens on the
// default button (Retry, No, OK, Fix, Yes), Esc is Cancel (Restore, No, OK); 44 px targets on a phone.
import { computed, ref, useId, watch } from 'vue'
import { ALERT_BUTTONS, ALERT_DEFAULT, type AlertAnswer, type AlertButton, type SheetAlert } from '../core/validation'
import { useMessages, type SheetMessages } from '../messages'
import DialogShell from './DialogShell.vue'

const props = withDefaults(
  defineProps<{
    /** The alert to show; null shows nothing. */
    alert: SheetAlert | null
    messages?: Partial<SheetMessages>
  }>(),
  { messages: undefined },
)
const emit = defineEmits<{
  /** The person's answer: `accept` (Yes, OK), `retry` (Retry, No) or `cancel`. */
  answer: [answer: AlertAnswer]
}>()

const m = useMessages(() => props.messages)
const id = useId()
const buttons = computed(() => (props.alert ? ALERT_BUTTONS[props.alert.severity] : []))
const label = (b: AlertButton): string =>
  ({ retry: m.value.retry, cancel: m.value.cancel, yes: m.value.validationYes, no: m.value.validationNo, ok: m.value.ok, fix: m.value.formulaFix, revert: m.value.formulaRevert })[b]
/** The icon's style: the sheet's own questions and messages show a warning's. */
const style = computed(() => {
  const severity = props.alert?.severity ?? 'stop'
  return severity === 'stop' || severity === 'information' ? severity : 'warning'
})

/** Each alert is drawn anew, so that it opens on its default button, as Excel's. */
const serial = ref(0)
watch(
  () => props.alert,
  () => serial.value++,
)
</script>

<template>
  <DialogShell
    v-if="alert"
    :key="serial"
    class="xs-va"
    role="alertdialog"
    :outside="false"
    :focus="`[data-xs-alert-button='${ALERT_DEFAULT[alert.severity]}']`"
    :aria-labelledby="`${id}-title`"
    :aria-describedby="`${id}-text`"
    :data-xs-alert="alert.severity"
    @keydown.stop
    @close="emit('answer', 'cancel')"
  >
    <header :id="`${id}-title`" class="xs-modal-title xs-va-title">{{ alert.title || m.validationTitle }}</header>
    <div class="xs-va-body">
      <svg class="xs-va-icon" viewBox="0 0 32 32" role="img" :aria-label="m.validationStyle(style)">
        <template v-if="style === 'stop'">
          <circle cx="16" cy="16" r="15" fill="var(--xs-alert-stop)" />
          <path d="M11 11l10 10M21 11 11 21" stroke="#fff" stroke-width="2.6" stroke-linecap="round" />
        </template>
        <template v-else-if="style === 'warning'">
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
        <p v-else-if="alert.severity === 'formula'" data-xs-alert-question>{{ m.formulaQuestion }}</p>
      </div>
    </div>
    <footer class="xs-va-foot">
      <button
        v-for="b in buttons"
        :key="b.button"
        type="button"
        :class="['xs-button', 'xs-va-btn', { 'xs-button-main': b.button === ALERT_DEFAULT[alert.severity] }]"
        :data-xs-alert-button="b.button"
        @click="emit('answer', b.answer)"
      >
        {{ label(b.button) }}
      </button>
    </footer>
  </DialogShell>
</template>

<style src="../controls.css"></style>
<style src="../validation.css"></style>
