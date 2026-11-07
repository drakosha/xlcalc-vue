// A text constant's UTF-16 runs as escaped text and CSS, including an explicit reset of each whole run font.
import type { CellStyle, FontStyle, TextRun } from 'xlcalc/contract'
import { decorationCss, fontCss } from './cell-style'

export interface RichRun {
  text: string
  style: Record<string, string>
}

export function richRuns(text: string, value: unknown, runs: TextRun[] | undefined, base: { name: string; size: number } | null, style: CellStyle | undefined, color: string | undefined, override?: FontStyle, link = false): RichRun[] | undefined {
  if (!runs?.length || typeof value !== 'string' || !value || !text.includes(value)) return undefined
  const out: RichRun[] = []
  const plain = (s: string) => { if (s) out.push({ text: s, style: decorationCss(style, link) }) }
  let at = 0
  for (let start = text.indexOf(value); start >= 0; start = text.indexOf(value, at)) {
    plain(text.slice(at, start))
    let end = 0
    for (let i = 0; i < runs.length; i++) {
      const run = runs[i]!
      const next = runs[i + 1]?.start ?? value.length
      if (!Number.isInteger(run.start) || run.start < end || run.start < 0 || next <= run.start || next > value.length) continue
      plain(value.slice(end, run.start))
      const font = { ...base, ...run.font, ...override }
      out.push({ text: value.slice(run.start, next), style: {
        fontFamily: 'var(--xs-font)', fontSize: 'var(--xs-font-size)',
        ...fontCss({ font }, color),
        fontWeight: font.bold ? '700' : '400', fontStyle: font.italic ? 'italic' : 'normal',
        color: color ?? font.color ?? (link ? 'var(--xs-link)' : 'var(--xs-fg)'),
        textDecorationLine: 'none', textDecorationStyle: 'solid', ...decorationCss({ font }, link),
      } })
      end = next
    }
    plain(value.slice(end))
    at = start + value.length
  }
  plain(text.slice(at))
  return out
}
