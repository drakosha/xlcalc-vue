// The PivotTable, slicer and timeline parts as a ref: null until their chunk loads, which starts when a sheet or a toolbar needs it.
import { shallowRef, watch, type ShallowRef } from 'vue'

export type PivotUi = typeof import('./pivot-ui')

const ui = shallowRef<PivotUi | null>(null)
let loading: Promise<PivotUi> | null = null

/** The parts for a part of the sheet that never asks for them first (the grid's buttons and boxes). */
export const pivotUi: Readonly<ShallowRef<PivotUi | null>> = ui

/** Loads the chunk once; after a failed load the next call tries again. */
export function loadPivotUi(): Promise<PivotUi> {
  loading ??= import('./pivot-ui').then(
    (m) => (ui.value = m),
    (e: unknown) => {
      loading = null
      throw e
    },
  )
  return loading
}

/**
 * The parts, loaded as soon as one of `wants()` holds: what the user asked for meanwhile is kept in the sheet's state (a
 * dialog's draft, a menu) and drawn when they arrive. `failed` runs when they cannot be loaded, to drop what waits for them;
 * the next change of `wants()` with one holding tries again.
 */
export function usePivotUi(wants: () => readonly unknown[], failed?: () => void): Readonly<ShallowRef<PivotUi | null>> {
  watch(
    wants,
    (list) => {
      if (!ui.value && list.some(Boolean)) loadPivotUi().catch(() => failed?.())
    },
    { immediate: true },
  )
  return pivotUi
}
