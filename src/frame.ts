// What the sheet and the toolbar both detect about their viewport: a narrow width, touch as the pointer, and how borders are drawn.
import { onScopeDispose, ref, type Ref } from 'vue'

/** Below this width what needs room gives way (panels below the grid, dialogs on the full sheet); with a finger as the pointer it is a phone. */
export const NARROW_PX = 600

/** Whether a media query matches, following it until the scope ends. */
function useMedia(queryText: string): Ref<boolean> {
  const matches = ref(false)
  if (typeof window !== 'undefined' && typeof window.matchMedia === 'function') {
    const query = window.matchMedia(queryText)
    const update = () => (matches.value = query.matches)
    update()
    query.addEventListener?.('change', update)
    onScopeDispose(() => query.removeEventListener?.('change', update))
  }
  return matches
}

/** Whether the pointer is touch (`pointer: coarse`). */
export const useCoarsePointer = (): Ref<boolean> => useMedia('(pointer: coarse)')
/** Short touch windows need compact controls even when the phone is landscape. */
export const useShortViewport = (): Ref<boolean> => useMedia('(max-height: 500px)')

/** The visible window for fixed menus and dialogs, including the on-screen keyboard. */
export function useVisualViewport() {
  const bounds = ref({ left: 0, top: 0, width: 0, height: 0 })
  if (typeof window === 'undefined') return bounds
  const viewport = window.visualViewport
  const update = () => { bounds.value = { left: viewport?.offsetLeft ?? 0, top: viewport?.offsetTop ?? 0, width: viewport?.width ?? window.innerWidth, height: viewport?.height ?? window.innerHeight } }
  update()
  window.addEventListener('resize', update)
  viewport?.addEventListener('resize', update)
  viewport?.addEventListener('scroll', update)
  onScopeDispose(() => {
    window.removeEventListener('resize', update)
    viewport?.removeEventListener('resize', update)
    viewport?.removeEventListener('scroll', update)
  })
  return bounds
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
