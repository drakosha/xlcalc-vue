// Excel's ready-made conditional formats: the formats of the quick rules' "with" box, the data bars, color scales and icon
// sets of its galleries with their default thresholds, and the rule each quick dialog writes.
import type { BorderLine, ColorScaleFormat, ConditionalFormat, ConditionalStyle, DataBarFormat, IconSetFormat, IconSetName, TimePeriodFormat } from 'xlcalc/contract'
import type { ColorScalePreset, DataBarColor, QuickRule } from '../contract/conditional'

/** The formats of the quick dialogs' "with" box, in its order; Custom Format follows them. */
export type StylePreset = 'lightRedDarkRed' | 'yellowDarkYellow' | 'greenDarkGreen' | 'lightRed' | 'redText' | 'redBorder'
export const STYLE_PRESETS: readonly StylePreset[] = ['lightRedDarkRed', 'yellowDarkYellow', 'greenDarkGreen', 'lightRed', 'redText', 'redBorder']

const DARK_RED = '#9C0006'
const RED_LINE: BorderLine = { style: 'thin', color: DARK_RED }
/** The colors of Excel's Bad, Neutral and Good cell styles, which the presets use, as Excel writes them to a file's dxf. */
export const PRESET_STYLES: Readonly<Record<StylePreset, ConditionalStyle>> = {
  lightRedDarkRed: { font: { color: DARK_RED }, fill: { color: '#FFC7CE' } },
  yellowDarkYellow: { font: { color: '#9C5700' }, fill: { color: '#FFEB9C' } },
  greenDarkGreen: { font: { color: '#006100' }, fill: { color: '#C6EFCE' } },
  lightRed: { fill: { color: '#FFC7CE' } },
  redText: { font: { color: DARK_RED } },
  redBorder: { border: { top: RED_LINE, right: RED_LINE, bottom: RED_LINE, left: RED_LINE } },
}

/** The colors of Excel's data bar gallery, the same for its gradient and solid fills. */
export const DATA_BAR_COLORS: Readonly<Record<DataBarColor, string>> = {
  blue: '#638EC6',
  green: '#63C384',
  red: '#FF555A',
  orange: '#FFB628',
  lightBlue: '#008AEF',
  purple: '#D6007B',
}

/** Excel 2010's data bar: automatic bounds, red negative bars, a black axis; a gradient bar has a border of its own color. */
export function dataBarRule(ranges: string[], color: DataBarColor, fill: 'gradient' | 'solid' = 'gradient'): DataBarFormat {
  const c = DATA_BAR_COLORS[color]
  const bar: DataBarFormat = { type: 'dataBar', ranges, min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: c, negativeColor: '#FF0000', axisColor: '#000000' }
  return fill === 'gradient' ? { ...bar, border: c, negativeBorder: '#FF0000' } : { ...bar, gradient: false }
}

const GREEN = '#63BE7B'
const YELLOW = '#FFEB84'
const RED = '#F8696B'
const WHITE = '#FCFCFF'
const BLUE = '#5A8AC6'
const PALE_YELLOW = '#FFEF9C'
/** The color scales of Excel's gallery, their colors from the lowest number to the highest. */
export const COLOR_SCALES: Readonly<Record<ColorScalePreset, readonly string[]>> = {
  greenYellowRed: [RED, YELLOW, GREEN],
  redYellowGreen: [GREEN, YELLOW, RED],
  greenWhiteRed: [RED, WHITE, GREEN],
  redWhiteGreen: [GREEN, WHITE, RED],
  blueWhiteRed: [RED, WHITE, BLUE],
  redWhiteBlue: [BLUE, WHITE, RED],
  whiteRed: [RED, WHITE],
  redWhite: [WHITE, RED],
  greenWhite: [WHITE, GREEN],
  whiteGreen: [GREEN, WHITE],
  greenYellow: [PALE_YELLOW, GREEN],
  yellowGreen: [GREEN, PALE_YELLOW],
}

/** A gallery color scale: the lowest and the highest number, and between them the 50th percentile of a three-color one. */
export function colorScaleRule(ranges: string[], preset: ColorScalePreset): ColorScaleFormat {
  const colors = COLOR_SCALES[preset]
  const values = colors.length === 3 ? [{ type: 'min' as const }, { type: 'percentile' as const, value: '50' }, { type: 'max' as const }] : [{ type: 'min' as const }, { type: 'max' as const }]
  return { type: 'colorScale', ranges, points: values.map((value, i) => ({ value, color: colors[i]! })) }
}

/** The number of icons in a set, which its name begins with. */
export const iconCount = (set: IconSetName): number => Number(set[0])

/** Excel's thresholds for an icon set: equal shares in percent (0, 33, 67 for three icons; 0, 25, 50, 75 for four; 0, 20, 40, 60, 80 for five). */
export function iconSetRule(ranges: string[], icons: IconSetName): IconSetFormat {
  const n = iconCount(icons)
  return { type: 'iconSet', ranges, icons, points: Array.from({ length: n }, (_, i) => ({ value: { type: 'percent' as const, value: String(Math.round((i * 100) / n)) } })) }
}

/** What a quick dialog was given: its boxes' formulas (stored, in the locale), the text or its formula, the period, Duplicate or Unique, the rank. */
export interface QuickInput {
  formulas?: string[]
  text?: string
  textFormula?: string
  period?: TimePeriodFormat['period']
  unique?: boolean
  rank?: number
  style: ConditionalStyle
}

/** The rule a quick dialog writes over `ranges`. */
export function quickRuleOf(kind: QuickRule, ranges: string[], q: QuickInput): ConditionalFormat {
  const style = q.style
  switch (kind) {
    case 'greaterThan':
    case 'lessThan':
    case 'between':
      return { type: 'cellIs', ranges, operator: kind, formulas: q.formulas ?? [], style }
    case 'equalTo':
      return { type: 'cellIs', ranges, operator: 'equal', formulas: q.formulas ?? [], style }
    case 'textContains':
      return { type: 'containsText', ranges, text: q.text ?? '', ...(q.textFormula ? { formula: q.textFormula } : {}), style }
    case 'dateOccurring':
      return { type: 'timePeriod', ranges, period: q.period ?? 'yesterday', style }
    case 'duplicateValues':
      return { type: q.unique ? 'uniqueValues' : 'duplicateValues', ranges, style }
    case 'aboveAverage':
      return { type: 'aboveAverage', ranges, style }
    case 'belowAverage':
      return { type: 'aboveAverage', ranges, below: true, style }
    default:
      return {
        type: 'top10',
        ranges,
        rank: q.rank ?? 10,
        ...(kind === 'top10Percent' || kind === 'bottom10Percent' ? { percent: true } : {}),
        ...(kind === 'bottom10' || kind === 'bottom10Percent' ? { bottom: true } : {}),
        style,
      }
  }
}

/** The quick rules whose dialog has a rank box, with its largest value: 1,000 items or 100 percent. */
export const rankMax = (kind: QuickRule): number | null => (kind === 'top10' || kind === 'bottom10' ? 1000 : kind === 'top10Percent' || kind === 'bottom10Percent' ? 100 : null)
