// Icons as small inline SVGs (24 × 24 strokes), so the package needs no icon library.
import { h, type FunctionalComponent } from 'vue'

/** The attributes every icon's svg has; the toolbar's colour icons draw on them too. */
export const ICON_ATTRS = { class: 'xs-icon', viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 2, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true' }

/** An icon of strokes along `paths`; `extra` overrides an attribute (a heavier stroke, dashes). */
export const icon = (paths: string[], extra: Record<string, unknown> = {}): FunctionalComponent => {
  const c: FunctionalComponent = () => h('svg', { ...ICON_ATTRS, ...extra }, paths.map((d) => h('path', { d })))
  return c
}

export const IconBack = icon(['M19 12H5', 'm12 19-7-7 7-7'])
export const IconLink = icon(['M9 17H7A5 5 0 0 1 7 7h2', 'M15 7h2a5 5 0 1 1 0 10h-2', 'M8 12h8'])
export const IconLock = icon(['M7 11V7a5 5 0 0 1 10 0v4', 'M5 11h14v10H5z'])
export const IconLockOpen = icon(['M7 11V7a5 5 0 0 1 9.9-1', 'M5 11h14v10H5z'])
export const IconClose = icon(['M18 6 6 18', 'm6 6 12 12'])
export const IconCheck = icon(['M20 6 9 17l-5-5'])
/** Excel's fx: an italic f and an x. */
export const IconFx = icon(['M11 4.5a2.5 2.5 0 0 0-3.6 1.6L5.2 17.9a2.5 2.5 0 0 1-3.6 1.6', 'M4.5 10H10', 'm13.5 10 7 9', 'm20.5 10-7 9'])
export const IconRetry = icon(['M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8', 'M3 3v5h5'])
export const IconSpinner = icon(['M21 12a9 9 0 1 1-6.22-8.56'])
export const IconSearch = icon(['M11 3a8 8 0 1 0 0 16a8 8 0 1 0 0-16', 'm21 21-4.3-4.3'])
export const IconFilter = icon(['M22 3H2l8 9.46V19l4 2v-8.54L22 3z'])
export const IconFilterOff = icon(['M13.013 3H2l8 9.46V19l4 2v-8.54l.9-1.055', 'm22 3-5 5', 'm17 3 5 5'])
export const IconUp = icon(['m18 15-6-6-6 6'])
export const IconDown = icon(['m6 9 6 6 6-6'])
export const IconSortAsc = icon(['m3 8 4-4 4 4', 'M7 4v16', 'M11 12h4', 'M11 16h7', 'M11 20h10'])
export const IconSortDesc = icon(['m3 16 4 4 4-4', 'M7 20V4', 'M11 4h10', 'M11 8h7', 'M11 12h4'])
export const IconPlus = icon(['M12 5v14', 'M5 12h14'])
export const IconOptions = icon(['M21 4h-7', 'M10 4H3', 'M21 12h-9', 'M8 12H3', 'M21 20h-5', 'M12 20H3', 'M14 2v4', 'M8 10v4', 'M16 18v4'])
