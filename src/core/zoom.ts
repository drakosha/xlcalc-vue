// Excel's zoom of a sheet, 10% to 400%: each sheet keeps its own, as Excel's sheet views do. Ctrl+wheel steps by 15%, a
// pinch (a trackpad's or two fingers) zooms in proportion about its center, the status bar and the Zoom dialog set it, and
// Zoom to Selection fits the selection. The grid's content is drawn under CSS zoom; its geometry stays in unzoomed pixels.
import { computed, nextTick, onScopeDispose, ref, watch } from 'vue'
import { sameName } from 'xlcalc/a1'
import type { SheetInfo } from '../contract/source'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

export const MIN_ZOOM = 10
export const MAX_ZOOM = 400
/** A notch of Ctrl+wheel zooms by 15%. */
export const WHEEL_STEP = 15
/** The status bar's − and + go to the next multiple of 10%. */
export const BUTTON_STEP = 10
/** A wheel event this far (or in lines or pages) is a mouse notch; a trackpad's pinch sends small deltas. */
const NOTCH_PX = 50
/** A pinch on a trackpad arrives as a burst of wheel events: it ends this long after the last one. */
const PINCH_IDLE_MS = 250

/** A zoom in whole percent within Excel's limits. */
export const clampZoom = (pct: number): number => Math.round(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Number.isFinite(pct) ? pct : 100)))

/** The status bar's − and +: the next multiple of 10% below or above. */
export const stepZoom = (pct: number, dir: 1 | -1): number => clampZoom(dir > 0 ? Math.floor(pct / BUTTON_STEP) * BUTTON_STEP + BUTTON_STEP : Math.ceil(pct / BUTTON_STEP) * BUTTON_STEP - BUTTON_STEP)

/** Excel's zoom slider: its left half covers 10% to 100%, its right half 100% to 400%, so 100% is at the middle (`pos` 0 to 100). */
export const zoomOfSlider = (pos: number): number => clampZoom(pos <= 50 ? MIN_ZOOM + (pos / 50) * (100 - MIN_ZOOM) : 100 + ((pos - 50) / 50) * (MAX_ZOOM - 100))
export const sliderOfZoom = (pct: number): number => (pct <= 100 ? ((pct - MIN_ZOOM) / (100 - MIN_ZOOM)) * 50 : 50 + ((pct - 100) / (MAX_ZOOM - 100)) * 50)

interface ZoomDeps {
  /** The host's zooms by sheet name (`SheetView`'s `zoom`). */
  given: () => Record<string, number> | undefined
  /** The user changed a sheet's zoom: the `zoom` event. */
  changed: (sheet: string, zoom: number) => void
}

export function useZoom(s: SheetState, v: Viewport, sel: Selection, deps: ZoomDeps) {
  /** Zooms set in this session by sheet name: they take precedence over the host's until it gives another. */
  const chosen = new Map<string, number>()
  const percent = computed(() => Math.round(s.zoom.value * 100))
  /** The zoom last reported for each sheet, so a gesture reports its end once. */
  const reported = new Map<string, number>()

  /** The zoom a sheet opens with: the host's for it, otherwise the workbook's stored one (`SheetInfo.zoom`, xlsx's `zoomScale`), otherwise 100%. */
  function initialZoom(sheet: SheetInfo): number {
    const given = deps.given()
    const name = given && Object.keys(given).find((k) => sameName(k, sheet.name))
    return clampZoom(name !== undefined ? given![name]! : (sheet.zoom ?? 100))
  }
  const zoomOf = (sheet: SheetInfo): number => chosen.get(sheet.name) ?? initialZoom(sheet)

  /**
   * Draws the active sheet at `pct`; `focus` (a point of the scroller, in page pixels from its corner) stays over the same
   * cell, as under a pinch, otherwise the top-left cell stays. Returns the zoom set.
   */
  function apply(pct: number, focus?: { x: number; y: number }): number {
    const next = clampZoom(pct)
    const sheet = s.sheet.value
    if (!sheet) return next
    chosen.set(sheet.name, next)
    const old = s.zoom.value
    const z = next / 100
    if (z === old) return next
    let top = s.scrollTop.value
    let left = s.scrollLeft.value
    if (focus) {
      top += focus.y / old - focus.y / z
      left += focus.x / old - focus.x / z
    }
    s.zoom.value = z
    sel.measure()
    s.scrollTop.value = Math.max(0, top)
    s.scrollLeft.value = Math.max(0, left)
    // the scroller's extent changes with the drawn zoom: scrolled once it is drawn
    void nextTick(() => s.scrollGridTo(s.scrollTop.value, s.scrollLeft.value))
    return next
  }
  /** Reports the active sheet's zoom when it differs from what was last reported for it (at first, what it opened with). */
  function report(): void {
    const sheet = s.sheet.value
    if (!sheet) return
    const pct = percent.value
    if ((reported.get(sheet.name) ?? initialZoom(sheet)) === pct) return
    reported.set(sheet.name, pct)
    deps.changed(sheet.name, pct)
  }
  /** A zoom the user chose (the status bar, the Zoom dialog, a wheel notch): set and reported. */
  function choose(pct: number): number {
    const got = apply(pct)
    report()
    return got
  }
  /** Excel's Zoom to Selection: the zoom at which the selection fills the viewport (10% to 400%), with the selection scrolled to its top-left corner; `user` reports it. */
  function fitSelection(user = false): number {
    const r = sel.selRect.value
    const el = s.scroller.value
    const top = r ? s.indexOf(r.top) : undefined
    const bottom = r ? s.indexOf(r.bottom) : undefined
    if (!r || !el || top === undefined || bottom === undefined) return percent.value
    const box = v.boxOf(top, r.left, bottom, r.right)
    const w = el.clientWidth || s.viewW.value * s.zoom.value
    const h = el.clientHeight || s.viewH.value * s.zoom.value
    const fit = Math.floor(Math.min(w / (v.leadW.value + Math.max(1, box.width)), h / (v.headH.value + v.footH.value + Math.max(1, box.height))) * 100)
    const got = apply(fit)
    // frozen rows and columns are in view already
    if (top > v.frozenRows.value) s.scrollTop.value = Math.max(0, box.top - v.headH.value - v.frozenH.value)
    if (r.left > v.frozenCols.value) s.scrollLeft.value = Math.max(0, box.left - v.leadW.value - v.frozenW.value)
    void nextTick(() => s.scrollGridTo(s.scrollTop.value, s.scrollLeft.value))
    if (user) report()
    return got
  }

  // ── gestures: Ctrl+wheel and a trackpad's pinch (Chrome and Firefox send it as Ctrl+wheel, Safari as gesture events), two fingers on a touch screen ──
  /** The zoom of a pinch in progress, unrounded, so its small steps add up. */
  let fine: number | null = null
  let idle: ReturnType<typeof setTimeout> | null = null
  const pointIn = (clientX: number, clientY: number) => {
    const box = s.scroller.value?.getBoundingClientRect()
    return box ? { x: clientX - box.left, y: clientY - box.top } : undefined
  }
  function pinchTo(z: number, focus?: { x: number; y: number }): void {
    fine = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z))
    apply(fine, focus)
  }
  function endPinch(): void {
    fine = null
    if (idle) clearTimeout(idle)
    idle = null
    report()
  }
  function onWheel(e: WheelEvent): void {
    if (!e.ctrlKey || !e.deltaY || !s.sheet.value) return
    e.preventDefault()
    if (e.deltaMode !== 0 || Math.abs(e.deltaY) >= NOTCH_PX) {
      endPinch()
      return void choose(percent.value - Math.sign(e.deltaY) * WHEEL_STEP)
    }
    pinchTo((fine ?? percent.value) * Math.exp(-e.deltaY / 100), pointIn(e.clientX, e.clientY))
    if (idle) clearTimeout(idle)
    idle = setTimeout(endPinch, PINCH_IDLE_MS)
  }
  /** Two fingers: their distance and the zoom when they touched. */
  let touch: { dist: number; zoom: number } | null = null
  const spread = (t: TouchList) => Math.hypot(t[0]!.clientX - t[1]!.clientX, t[0]!.clientY - t[1]!.clientY)
  function onTouchStart(e: TouchEvent): void {
    if (e.touches.length !== 2 || !s.sheet.value) return
    touch = { dist: Math.max(1, spread(e.touches)), zoom: percent.value }
  }
  function onTouchMove(e: TouchEvent): void {
    if (!touch || e.touches.length !== 2) return
    // the page would zoom or scroll under the fingers instead
    if (e.cancelable) e.preventDefault()
    const [a, b] = [e.touches[0]!, e.touches[1]!]
    pinchTo((touch.zoom * spread(e.touches)) / touch.dist, pointIn((a.clientX + b.clientX) / 2, (a.clientY + b.clientY) / 2))
  }
  function onTouchEnd(e: TouchEvent): void {
    if (!touch || e.touches.length >= 2) return
    touch = null
    endPinch()
  }
  /** Safari's pinch on a trackpad (and alongside the touches on iOS, which the touch handlers already follow). */
  let gesture: number | null = null
  function onGesture(e: Event & { scale?: number; clientX?: number; clientY?: number }): void {
    if (!s.sheet.value) return
    e.preventDefault()
    if (touch) return
    if (e.type === 'gesturestart') gesture = percent.value
    else if (e.type === 'gestureend') {
      gesture = null
      endPinch()
    } else if (gesture !== null && e.scale) pinchTo(gesture * e.scale, e.clientX === undefined ? undefined : pointIn(e.clientX, e.clientY ?? 0))
  }
  onScopeDispose(() => idle && clearTimeout(idle))

  /** A sheet made active: drawn at its own zoom. */
  function enter(): void {
    const sheet = s.sheet.value
    s.zoom.value = sheet ? zoomOf(sheet) / 100 : 1
    sel.measure()
  }
  // The host's zooms: a sheet it gives a new zoom for takes it; the active one is drawn at it at once.
  watch(deps.given, (next, prev) => {
    for (const [name, pct] of Object.entries(next ?? {})) {
      if (prev && prev[name] === pct) continue
      for (const map of [chosen, reported]) for (const key of [...map.keys()]) if (sameName(key, name)) map.delete(key)
    }
    if (s.sheet.value) apply(zoomOf(s.sheet.value))
  })

  /** Excel's Zoom dialog (a click on the status bar's percentage). */
  const dialogOpen = ref(false)

  return { percent, apply, choose, fitSelection, enter, onWheel, onTouchStart, onTouchMove, onTouchEnd, onGesture, dialogOpen }
}

export type Zoom = ReturnType<typeof useZoom>
