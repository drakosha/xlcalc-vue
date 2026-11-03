// What the sheet and the toolbar both detect about their viewport: a phone width, touch as the pointer, and how borders are drawn.
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

/** The width a hairline border is drawn at, in page pixels, measured: a CSS pixel in Chromium, a device pixel in WebKit; 1 without a layout. */
function measureHairline(): number {
  const d = document.createElement('div')
  d.style.cssText = 'all:initial;position:absolute;visibility:hidden;width:0;height:0;border-left:0.01px solid'
  document.body.appendChild(d)
  const w = d.getBoundingClientRect().width
  d.remove()
  // A whole fraction of a page pixel (layout rounds a third to 1/64ths); anything else is the host's styles: one pixel.
  return w > 0 && w <= 1 ? 1 / Math.round(1 / w) : 1
}

/** Borders are drawn in whole steps of this many page pixels (`measureHairline`), measured again when the pixel ratio changes. */
export function useHairline(): Ref<number> {
  const step = ref(1)
  if (typeof window !== 'undefined' && typeof window.matchMedia === 'function') {
    let query: MediaQueryList | null = null
    const update = () => {
      query?.removeEventListener?.('change', update)
      step.value = measureHairline()
      query = window.matchMedia(`(resolution: ${window.devicePixelRatio}dppx)`)
      query.addEventListener?.('change', update)
    }
    update()
    onScopeDispose(() => query?.removeEventListener?.('change', update))
  }
  return step
}
