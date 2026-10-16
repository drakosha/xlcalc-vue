// What the boxes over the grid share: a list or a tip under its field (over it where the view has no room below), kept
// inside the view sideways, the list's highlighted item kept in view, and a press outside that closes one.
import { nextTick, onBeforeUnmount, watch, type Ref, type WatchSource } from 'vue'

/** A box, in the coordinates of the element the popups are placed in. */
export interface PopupRect {
  top: number
  left: number
  width: number
  height: number
}

/** An item of Excel's lists (AutoComplete, a validation list) is 18 px; 44 under a finger, as --xs-touch-target in the CSS. */
export const ITEM_PX = 18
export const TOUCH_PX = 44

/** A list's height: up to `rows` items and its 1 px border above and below. */
export const listHeight = (count: number, rows: number, touch: boolean): number => Math.min(count, rows) * (touch ? TOUCH_PX : ITEM_PX) + 2

/**
 * The top of a box `height` high under `field` (`gap` px from it), or over it when the view has no room below and room
 * above; under it without a view or with room neither way.
 */
export function placeBelowOrAbove(field: PopupRect, height: number, view: PopupRect | undefined, gap = 0): number {
  const below = field.top + field.height + gap
  const above = field.top - gap - height
  return view && below + height > view.top + view.height && above >= view.top ? above : below
}

/** A box `width` wide from `left`, moved left to end at the view's right edge, but not past its left one. */
export const clampLeft = (left: number, width: number, view: PopupRect | undefined): number =>
  view ? Math.max(view.left, Math.min(left, view.left + view.width - width)) : left

/** Scrolls the list (the offset parent of its items) by its own scroll alone, so the page and its containers stay, until item `i` is in view. */
export function scrollToItem(list: HTMLElement | null | undefined, i: number): void {
  const item = list?.children[i] as HTMLElement | undefined
  if (!list || !item) return
  if (item.offsetTop < list.scrollTop) list.scrollTop = item.offsetTop
  else if (item.offsetTop + item.offsetHeight > list.scrollTop + list.clientHeight) list.scrollTop = item.offsetTop + item.offsetHeight - list.clientHeight
}

/** Keeps a list's highlighted item in view as the keys move it: after each change of `on`, once the list is drawn. */
export function keepActiveInView(list: Ref<HTMLElement | null>, on: WatchSource<unknown>, active: () => number): void {
  watch(on, async () => {
    await nextTick()
    scrollToItem(list.value, active())
  })
}

/**
 * Calls `onOutside` on a press outside `inside` (an element, or a selector the pressed element's ancestors match) while
 * `active` says so; the press is caught before the page's own handlers see it.
 */
export function useOutsidePress(inside: Ref<HTMLElement | null> | string, onOutside: () => void, active: () => boolean): void {
  function onDown(e: PointerEvent): void {
    const target = e.target as Element | null
    const within = typeof inside === 'string' ? !!target?.closest?.(inside) : !!inside.value?.contains(target)
    if (!within) onOutside()
  }
  const off = () => document.removeEventListener('pointerdown', onDown, true)
  // no document on a server's render: nothing to listen to
  watch(
    active,
    (on) => {
      if (typeof document === 'undefined') return
      if (on) document.addEventListener('pointerdown', onDown, true)
      else off()
    },
    { immediate: true },
  )
  onBeforeUnmount(off)
}
