// The PivotTable, slicer and timeline parts as a ref: null until their chunk loads, which starts when a sheet or a toolbar needs it.
import { shallowRef, watch, type ShallowRef } from 'vue'

export type PivotUi = typeof import('./pivot-ui')

const ui = shallowRef<PivotUi | null>(null)
let loading: Promise<PivotUi> | null = null
/** A load has failed: only a new request of the user's tries again, not what the sheet shows. */
let failedOnce = false

/** The parts for a part of the sheet that never asks for them first (the grid's buttons and boxes). */
export const pivotUi: Readonly<ShallowRef<PivotUi | null>> = ui

/** Loads the chunk once; after a failed load the next call tries again. */
export function loadPivotUi(): Promise<PivotUi> {
  loading ??= import('./pivot-ui').then(
    (m) => (ui.value = m),
    (e: unknown) => {
      loading = null
      failedOnce = true
      throw e
    },
  )
  return loading
}

/**
 * The parts, loaded as soon as `shown()` holds (the sheet has a PivotTable, a slicer or a timeline) or one of `requests()`
 * appears (a dialog, a menu, a draft the user asked for): a request made meanwhile stays in the sheet's state and is drawn
 * when they arrive. When they cannot be loaded, `failed` drops the requests that waited; after that only a new request
 * tries again, so moving about a PivotTable neither loads them again nor moves the focus.
 */
export function usePivotUi(shown: () => boolean, requests: () => readonly unknown[] = () => [], failed?: () => void): Readonly<ShallowRef<PivotUi | null>> {
  const load = () => void loadPivotUi().catch(() => failed?.())
  watch(
    shown,
    (on) => {
      if (on && !ui.value && !failedOnce) load()
    },
    { immediate: true },
  )
  watch(
    requests,
    (now, before) => {
      if (!ui.value && now.some((r, i) => r && r !== before?.[i])) load()
    },
    { immediate: true },
  )
  return pivotUi
}
