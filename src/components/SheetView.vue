<script setup lang="ts">
// A workbook sheet over a SheetSource: the formula bar, the grid with windows, styles, editing and zoom, the Trace Precedents panel, the Watch Window, the tabs, and messages.
import { getCurrentInstance, provide } from 'vue'
import type { ListSort, SheetFailure, SheetLinkEvent, SheetViewApi, SheetViewProps, SheetViewState, SheetWatch, SheetZoomEvent } from '../contract/view'
import { useSheet, type SheetEmit } from '../core/sheet'
import { sheetContextKey } from './context'
import { hasListener } from './dom'
import CfDialogs from './CfDialogs.vue'
import ExplainPanel from './ExplainPanel.vue'
import FilterPopover from './FilterPopover.vue'
import FindBar from './FindBar.vue'
import FormatDialog from './FormatDialog.vue'
import GoToDialog from './GoToDialog.vue'
import SymbolDialog from './SymbolDialog.vue'
import InsertFunction from './InsertFunction.vue'
import PasteSpecialDialog from './PasteSpecialDialog.vue'
import LinkDialog from './LinkDialog.vue'
import ValidationAlert from './ValidationAlert.vue'
import WatchWindow from './WatchWindow.vue'
import ZoomDialog from './ZoomDialog.vue'
import ValidationDialog from './ValidationDialog.vue'
import FormulaBar from './FormulaBar.vue'
import NameDialogs from './NameDialogs.vue'
import SheetGrid from './SheetGrid.vue'
import SheetMenu from './SheetMenu.vue'
import SheetTabs from './SheetTabs.vue'
import SpellingDialog from './SpellingDialog.vue'
import TableDialog from './TableDialog.vue'
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
  zoom: undefined,
  watches: undefined,
  noteAuthor: undefined,
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
  /** The user changed a sheet's zoom; the host keeps it (the `zoom` prop) to open the sheet at it again. */
  zoom: [change: SheetZoomEvent]
  /** The user added or deleted watches in the Watch Window: the whole list, which the host keeps (the `watches` prop). */
  watches: [list: SheetWatch[]]
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
      <WatchWindow v-if="ctx.watches.open.value" />
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
      <GoToDialog v-if="ctx.goto.open.value" />
      <SymbolDialog v-if="ctx.symbolOpen.value" />
      <PasteSpecialDialog v-if="ctx.clip.specialOpen.value" />
      <InsertFunction v-if="ctx.insertFunctionOpen.value" @pick="ctx.closeInsertFunction" @close="ctx.closeInsertFunction()" />
      <NameDialogs v-if="ctx.names.available.value" />
      <ZoomDialog v-if="ctx.zoomer.dialogOpen.value" />
      <TableDialog />
      <SpellingDialog />
      <LinkDialog />
      <ValidationDialog />
      <CfDialogs />
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
