// The app bar's icons: a new book, open, save as, settings and close, as small inline SVGs.
import { h, type FunctionalComponent } from 'vue'

const icon = (paths: string[]): FunctionalComponent => {
  const c: FunctionalComponent = () =>
    h(
      'svg',
      { class: 'spa-icon', viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 1.8, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true' },
      paths.map((d) => h('path', { d })),
    )
  return c
}

export const IconNew = icon(['M14 3H6.5A1.5 1.5 0 0 0 5 4.5v15A1.5 1.5 0 0 0 6.5 21h11a1.5 1.5 0 0 0 1.5-1.5V8z', 'M14 3v5h5', 'M12 11v6', 'M9 14h6'])
export const IconOpen = icon(['M3 7.5A1.5 1.5 0 0 1 4.5 6H9l2 2h8.5A1.5 1.5 0 0 1 21 9.5v8a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 17.5z'])
export const IconSave = icon(['M12 4v11', 'm7 10 5 5 5-5', 'M5 20h14'])
export const IconSettings = icon(['M4 7h10', 'M18 7h2', 'M4 17h2', 'M10 17h10', 'M16 5v4', 'M8 15v4'])
export const IconClose = icon(['M18 6 6 18', 'm6 6 12 12'])
