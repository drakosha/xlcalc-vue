<script setup lang="ts">
// A workbook sheet over a SheetSource: the formula bar, the grid with windows, styles and editing, the Trace Precedents panel, the tabs, and messages.
import { getCurrentInstance, provide } from 'vue'
import type { ListSort, SheetFailure, SheetLinkEvent, SheetViewApi, SheetViewProps, SheetViewState } from '../contract/view'
import { useSheet, type SheetEmit } from '../core/sheet'
import { sheetContextKey } from './context'
import { hasListener } from './dom'
import ExplainPanel from './ExplainPanel.vue'
import FilterPopover from './FilterPopover.vue'
import FindBar from './FindBar.vue'
import FormatDialog from './FormatDialog.vue'
import ValidationAlert from './ValidationAlert.vue'
import FormulaBar from './FormulaBar.vue'
import NameDialogs from './NameDialogs.vue'
import SheetGrid from './SheetGrid.vue'
import SheetMenu from './SheetMenu.vue'
import SheetTabs from './SheetTabs.vue'
import { IconClose, IconRetry, IconSpinner } from './icons'

const props = withDefaults(defineProps<SheetViewProps>(), {
  messages: undefined,
  storageKey: undefined,
  focus: null,
  cellUrl: undefined,
  readOnly: false,
  formulaBar: true,
  locale: undefined,
  hideLoneTab: false,
  edgeRestoresWidth: false,
  sort: null,
  openRows: false,
  hostLinks: false,
})
const emit = defineEmits<{
  state: [state: SheetViewState | null]
  openLink: [link: SheetLinkEvent]
  /** List mode: a sortable column's title was clicked; the host sorts the list and sends `reload`. */
  sort: [sort: ListSort]
  /** List mode: a drill-down cell was clicked or Enter was pressed on it. */
  cellDrill: [cell: { sheet: string; cell: string }]
  /** List mode (`openRows`): a row was opened from its first column or by Enter. */
  rowOpen: [row: { sheet: string; row: number }]
}>()
defineSlots<{
  /** The workbook cannot be displayed: the source's error and a retry function. */
  fallback?: (scope: { failure: SheetFailure; retry: () => void }) => unknown
}>()

const instance = getCurrentInstance()
/** Whether the host opens links (`openLink`): otherwise a web link opens as an ordinary link. */
const linksToHost = () => props.hostLinks || hasListener(instance, 'onOpenLink')
const ctx = useSheet(props, emit as SheetEmit, linksToHost)
provide(sheetContextKey, ctx)
defineExpose<SheetViewApi>(ctx.api)
const m = ctx.messages
const failure = ctx.failure
const message = ctx.message
const notice = ctx.w.notice
const wait = ctx.w.wait
</script>

<template>
  <div class="xs-sheet" :class="{ 'xs-narrow': ctx.v.narrow.value, 'xs-touch': ctx.v.touch.value }">
    <template v-if="failure">
      <slot name="fallback" :failure="failure" :retry="ctx.retry">
        <div class="xs-fallback" role="alert">
          <p>{{ failure.text }}</p>
          <button type="button" class="xs-button" @click="ctx.retry()">
            <IconRetry />
            {{ m.retry }}
          </button>
        </div>
      </slot>
    </template>
    <template v-else>
      <FormulaBar v-if="formulaBar" />
      <div v-if="notice" class="xs-notice">{{ notice }}</div>
      <FindBar v-if="ctx.find.isOpen.value" />
      <div class="xs-main">
        <div class="xs-body">
          <SheetGrid />
          <!-- A workbook being recalculated: the sheet waits over the grid with a stopwatch and requests again every second. -->
          <div v-if="wait" class="xs-wait" role="status" aria-live="polite" data-xs-wait>
            <IconSpinner class="xs-spin xs-wait-icon" />
            <p class="xs-wait-title">{{ m.waiting }}</p>
            <p class="xs-wait-seconds">{{ m.waitSeconds(ctx.w.waitSeconds.value) }}</p>
            <p class="xs-muted">{{ m.waitHint }}</p>
          </div>
        </div>
        <ExplainPanel v-if="ctx.explain.open.value" />
      </div>
      <SheetTabs />
      <FilterPopover />
      <SheetMenu />
      <FormatDialog
        :open="ctx.cmd.formatOpen.value"
        :state="ctx.viewState.value"
        :messages="m"
        :locale="locale"
        @apply="(patches) => void ctx.cmd.format(patches)"
        @close="ctx.cmd.closeFormat()"
      />
      <NameDialogs v-if="ctx.names.available.value" />
      <ValidationAlert :alert="ctx.alerts.alert.value" :messages="m" @answer="ctx.alerts.answer" />
      <div v-if="message" class="xs-message" role="status" aria-live="polite">
        <span>{{ message }}</span>
        <button type="button" class="xs-tool" :title="m.close" @click="ctx.dismiss()">
          <IconClose />
        </button>
      </div>
    </template>
  </div>
</template>

<style src="../style.css"></style>
