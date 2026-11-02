// A large toolbar button's label under its icon, as Excel's ribbon draws it: one or two lines, a menu's arrow after the last.
import { h, type FunctionalComponent } from 'vue'
import { largeLabel } from '../toolbar'
import { IconDown } from './icons'

/** The label of `name` (its key left out); with `arrow` the menu's arrow ends it. Hidden from the accessible name, which the button's `aria-label` gives. */
export const ToolbarLabel: FunctionalComponent<{ name: string; arrow?: boolean }> = ({ name, arrow }) => {
  const lines = largeLabel(name, !!arrow)
  return h(
    'span',
    { class: 'xs-tb-label', 'aria-hidden': 'true' },
    lines.map((text, i) => h('span', { class: 'xs-tb-label-line' }, arrow && i === lines.length - 1 ? [text, h(IconDown, { class: 'xs-tb-label-down' })] : text)),
  )
}
// `arrow` as a boolean prop, so that a bare `arrow` attribute turns it on
ToolbarLabel.props = { name: { type: String, required: true }, arrow: Boolean }
