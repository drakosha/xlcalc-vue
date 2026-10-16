// Keys of Excel's lists of items (AutoComplete's functions, a validation list): how far each moves the highlight, and where it lands.

/** How far a key moves a list's highlight: an item, a page of `page` items, or the whole list (Home, End); null for another key. */
export function listStep(key: string, count: number, page: number): number | null {
  switch (key) {
    case 'ArrowDown':
      return 1
    case 'ArrowUp':
      return -1
    case 'PageDown':
      return page
    case 'PageUp':
      return -page
    case 'Home':
      return -count
    case 'End':
      return count
    default:
      return null
  }
}

/** The item `by` steps from `index` among `count` (at least one), kept within the list; from none (-1) down goes to the first, up to the last. */
export function stepIndex(index: number, by: number, count: number): number {
  const from = index < 0 ? (by > 0 ? -1 : count) : index
  return Math.min(count - 1, Math.max(0, from + by))
}
