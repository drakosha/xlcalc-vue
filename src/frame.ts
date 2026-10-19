// What the sheet and the toolbar both detect about their viewport: a phone width, and touch as the pointer.
import { onScopeDispose, ref, type Ref } from 'vue'

/** Below this width a viewport counts as a phone's: narrower row numbers, the toolbar's buttons in a menu. */
export const NARROW_PX = 600

/** Whether the pointer is touch (`pointer: coarse`), following the media query until the scope ends. */
export function useCoarsePointer(): Ref<boolean> {
  const coarse = ref(false)
  if (typeof window !== 'undefined' && typeof window.matchMedia === 'function') {
    const query = window.matchMedia('(pointer: coarse)')
    const update = () => (coarse.value = query.matches)
    update()
    query.addEventListener?.('change', update)
    onScopeDispose(() => query.removeEventListener?.('change', update))
  }
  return coarse
}
