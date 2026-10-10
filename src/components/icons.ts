// Icons as small inline SVGs (24 × 24 strokes, the look of the the app's), so the package needs no icon library.
import { h, type FunctionalComponent } from 'vue'

const icon = (paths: string[]): FunctionalComponent => {
  const c: FunctionalComponent = () =>
    h(
      'svg',
      { class: 'xs-icon', viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 2, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true' },
      paths.map((d) => h('path', { d })),
    )
  return c
}

export const IconBack = icon(['M9 14 4 9l5-5', 'M4 9h10.5a5.5 5.5 0 0 1 0 11H11'])
export const IconLink = icon(['M9 17H7A5 5 0 0 1 7 7h2', 'M15 7h2a5 5 0 1 1 0 10h-2', 'M8 12h8'])
export const IconLock = icon(['M7 11V7a5 5 0 0 1 10 0v4', 'M5 11h14v10H5z'])
export const IconLockOpen = icon(['M7 11V7a5 5 0 0 1 9.9-1', 'M5 11h14v10H5z'])
export const IconClose = icon(['M18 6 6 18', 'm6 6 12 12'])
export const IconRetry = icon(['M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8', 'M3 3v5h5'])
export const IconSpinner = icon(['M21 12a9 9 0 1 1-6.22-8.56'])
