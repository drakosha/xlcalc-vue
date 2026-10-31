// The PivotTable Fields pane's state over a source and a sheet: the PivotTable under the active cell read through the source's
// `pivotTable` (again on its `reload` and `sheets` events), its areas, Defer Layout Update with the layout waiting for Update,
// and each change sent as one `editPivotTable` through the sheet's `structure`, one step of undo.
import { computed, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import type { SheetPivotTable } from '../contract/source'
import type { PivotDialogRequest, PivotFieldsPaneProps } from '../contract/view'
import { areasOf, patchOf, type PivotAreas } from '../core/pivot-layout'

export function usePivotPane(props: PivotFieldsPaneProps) {
  const definition = shallowRef<SheetPivotTable | null>(null)
  /** The layout shown while it waits: for Update with Defer Layout Update on, or for the source's answer. */
  const pending = shallowRef<PivotAreas | null>(null)
  const defer = ref(false)
  const failed = ref(false)
  /** The PivotTable shown: the sheet and name of the state's, while the pane is on. */
  const target = computed(() => (props.state?.fieldList && props.state.pivot ? { sheet: props.state.sheet, pivot: props.state.pivot } : null))
  const areas = computed<PivotAreas | null>(() => pending.value ?? (definition.value ? areasOf(definition.value.table) : null))
  const fields = computed(() => definition.value?.cache.fields ?? [])
  /** Changes are sent through the sheet, which an editable sheet with `structure` takes. */
  const editable = computed(() => !!props.sheet && !!props.source.structure)

  let reading: AbortController | null = null
  async function read(): Promise<void> {
    reading?.abort()
    const t = target.value
    const call = props.source.pivotTable
    if (!t || !call) {
      definition.value = null
      return
    }
    const ctl = (reading = new AbortController())
    try {
      const got = await call.call(props.source, t.sheet, t.pivot, ctl.signal)
      if (ctl.signal.aborted) return
      definition.value = got
      failed.value = false
    } catch {
      if (!ctl.signal.aborted) failed.value = true
    }
  }
  // another PivotTable: its own layout, Defer Layout Update off as Excel shows a PivotTable's own setting
  watch(
    () => (target.value ? `${target.value.sheet}\n${target.value.pivot}` : ''),
    () => {
      pending.value = null
      defer.value = false
      void read()
    },
    { immediate: true },
  )
  const unsubscribe = props.source.subscribe?.((e) => {
    if (e.kind === 'reload' || e.kind === 'sheets') void read()
  })
  onBeforeUnmount(() => {
    reading?.abort()
    unsubscribe?.()
  })

  let sending = Promise.resolve()
  /** The layout sent as one `editPivotTable`; shown until the source answers, then read again. */
  function send(next: PivotAreas): Promise<void> {
    sending = sending.then(async () => {
      const d = definition.value
      const t = target.value
      const set = d && patchOf(d.table, next)
      if (!d || !t || !set || !props.sheet) return
      pending.value = next
      try {
        await props.sheet.structure({ op: 'editPivotTable', sheet: t.sheet, pivot: t.pivot, set })
        await read()
      } finally {
        pending.value = null
      }
    })
    return sending
  }

  /** A change of the areas: sent at once, or kept for Update with Defer Layout Update on. */
  function apply(next: PivotAreas | null): void {
    if (!next || !editable.value) return
    if (defer.value) pending.value = next
    else void send(next)
  }
  /** Update: the layout kept since Defer Layout Update was checked, in one step of undo. */
  function update(): Promise<void> {
    const next = pending.value
    return next ? send(next) : Promise.resolve()
  }
  /** Unchecking Defer Layout Update lays out what waited, as Excel does. */
  function setDefer(on: boolean): void {
    defer.value = on
    if (!on) void update()
  }
  /** Whether an Update waits. */
  const waiting = computed(() => defer.value && !!pending.value && !!definition.value && !!patchOf(definition.value.table, pending.value))

  /** Asks the sheet for a dialog over a field, as the pane's menus do. */
  function dialog(request: Omit<PivotDialogRequest, 'sheet' | 'pivot'>): void {
    const t = target.value
    if (t) props.sheet?.pivotDialog({ ...request, sheet: t.sheet, pivot: t.pivot })
  }

  return { definition, areas, fields, target, editable, defer, waiting, failed, apply, update, setDefer, dialog }
}

export type PivotPane = ReturnType<typeof usePivotPane>
