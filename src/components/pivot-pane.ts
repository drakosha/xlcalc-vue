// The PivotTable Fields pane's state over a source and a sheet: the PivotTable under the active cell read through the source's
// `pivotTable` (again on its `reload` and `sheets` events), its areas, Defer Layout Update with the layout waiting for Update,
// and each change sent as one `editPivotTable` through the sheet's `structure`, one step of undo.
import { computed, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import type { PivotDataField } from 'xlcalc/contract'
import type { SheetPivotTable } from '../contract/source'
import type { PivotDialogRequest, PivotFieldsPaneProps } from '../contract/view'
import { areasOf, patchOf, type PivotAreas } from '../core/pivot-layout'
import { allows } from '../protection'

export function usePivotPane(props: PivotFieldsPaneProps) {
  const definition = shallowRef<SheetPivotTable | null>(null)
  /** The last layout sent or queued and not yet answered: what the next change builds on, and what shows meanwhile. */
  const requested = shallowRef<PivotAreas | null>(null)
  /** The layout kept for Update while Defer Layout Update is on. */
  const draft = shallowRef<PivotAreas | null>(null)
  const defer = ref(false)
  const failed = ref(false)
  /** The PivotTable shown: the sheet and name of the state's, while the pane is on. */
  const target = computed(() => (props.state?.fieldList && props.state.pivot ? { sheet: props.state.sheet, pivot: props.state.pivot } : null))
  /** The PivotTable as the source holds it. */
  const saved = computed<PivotAreas | null>(() => (definition.value ? areasOf(definition.value.table) : null))
  /** What the pane shows, and what a change builds on: the draft, else the layout on its way, else the PivotTable's. */
  const areas = computed<PivotAreas | null>(() => draft.value ?? requested.value ?? saved.value)
  const cache = computed(() => definition.value?.cache ?? null)
  const fields = computed(() => cache.value?.fields ?? [])
  /**
   * Changes go through the sheet where it takes them: an editable sheet with `structure` (Insert PivotTable is available
   * there), whose protection allows Use PivotTable and PivotChart.
   */
  const editable = computed(() => !!props.sheet && !!props.state?.canInsertPivot && allows(props.state.protection, 'pivotTables'))

  /** Changes sent and not yet answered. */
  let inFlight = 0
  let reading: AbortController | null = null
  /**
   * Reads the PivotTable again. On a change from elsewhere (`outside`: an undo, a Refresh) with nothing of the pane's on its
   * way, a draft built on the old layout is dropped, as it no longer describes a change of what is there.
   */
  async function read(outside = false): Promise<void> {
    reading?.abort()
    const t = target.value
    const call = props.source.pivotTable
    if (!t || !call) {
      definition.value = null
      return
    }
    const ctl = (reading = new AbortController())
    try {
      const before = saved.value
      const got = await call.call(props.source, t.sheet, t.pivot, ctl.signal)
      if (ctl.signal.aborted) return
      definition.value = got
      failed.value = false
      if (outside && !inFlight && draft.value && JSON.stringify(before) !== JSON.stringify(saved.value)) draft.value = null
    } catch {
      if (!ctl.signal.aborted) failed.value = true
    }
  }
  // another PivotTable: its own layout, Defer Layout Update off as Excel shows a PivotTable's own setting
  watch(
    () => (target.value ? `${target.value.sheet}\n${target.value.pivot}` : ''),
    () => {
      requested.value = null
      draft.value = null
      defer.value = false
      void read()
    },
    { immediate: true },
  )
  const unsubscribe = props.source.subscribe?.((e) => {
    if (e.kind === 'reload' || e.kind === 'sheets') void read(true)
  })
  onBeforeUnmount(() => {
    reading?.abort()
    unsubscribe?.()
  })

  let sending = Promise.resolve()
  /**
   * The layout sent as one `editPivotTable`, after the changes before it: its patch from the PivotTable as read after them.
   * It shows until the source answers; the last one answered, the PivotTable shows as read again, refused or not.
   */
  function send(next: PivotAreas): Promise<void> {
    requested.value = next
    inFlight++
    sending = sending
      .then(async () => {
        const d = definition.value
        const t = target.value
        const set = d && patchOf(d.table, next)
        if (!d || !t || !set || !props.sheet) return
        try {
          // a field of dates put on the rows or columns is grouped as Excel's field list groups it (xlcalc's `groupDates`)
          await props.sheet.structure({ op: 'editPivotTable', sheet: t.sheet, pivot: t.pivot, set, groupDates: true })
        } finally {
          await read()
        }
      })
      .catch(() => {})
      .finally(() => {
        inFlight--
        if (requested.value === next) requested.value = null
      })
    return sending
  }

  /** A change of the areas: sent at once, or kept for Update with Defer Layout Update on. */
  function apply(next: PivotAreas | null): void {
    if (!next || !editable.value) return
    if (defer.value) draft.value = next
    else void send(next)
  }
  /** Update: the layout kept since Defer Layout Update was checked, in one step of undo; a change made meanwhile is a new draft. */
  function update(): Promise<void> {
    const next = draft.value
    draft.value = null
    return next ? send(next) : Promise.resolve()
  }
  /** Unchecking Defer Layout Update lays out what waited, as Excel does. */
  function setDefer(on: boolean): void {
    defer.value = on
    if (!on) void update()
  }
  /** Whether an Update waits: a draft that differs from the PivotTable. */
  const waiting = computed(() => !!draft.value && !!definition.value && !!patchOf(definition.value.table, draft.value))

  /**
   * A value field's index in the PivotTable as the source holds it, which a host's Value Field Settings reads: the value field
   * shown at `i`, found among the saved ones; null for one that only the draft or a change on its way has.
   */
  function savedValue(v: PivotDataField): number | null {
    const at = definition.value?.table.values?.findIndex((x) => JSON.stringify(x) === JSON.stringify(v)) ?? -1
    return at < 0 ? null : at
  }

  /** Asks the sheet for a dialog over a field, as the pane's menus do. */
  function dialog(request: Omit<PivotDialogRequest, 'sheet' | 'pivot'>): void {
    const t = target.value
    if (t) props.sheet?.pivotDialog({ ...request, sheet: t.sheet, pivot: t.pivot })
  }

  return { definition, areas, cache, fields, target, editable, defer, waiting, failed, apply, update, setDefer, savedValue, dialog }
}

export type PivotPane = ReturnType<typeof usePivotPane>
