// The sheet's look beside its cells: its tab's color.

/** Whether a `#RRGGBB` color is dark, by its perceived brightness: a tab filled with it takes light text. */
export function darkColour(hex: string): boolean {
  const n = Number.parseInt(hex.slice(1, 7), 16)
  if (Number.isNaN(n)) return false
  const r = (n >> 16) & 255
  const g = (n >> 8) & 255
  const b = n & 255
  return (r * 299 + g * 587 + b * 114) / 1000 < 128
}
