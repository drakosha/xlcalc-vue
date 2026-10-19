// The icons of Excel's 20 icon sets as small inline SVGs (16 × 16), in the colors Excel for Mac 16 draws them, by set and number.
import { h, type FunctionalComponent } from 'vue'
import type { Icon, IconSetName } from 'xlcalc/contract'

/** One element of an icon: its tag and attributes. */
type Shape = [string, Record<string, string | number>]
/** A fill and its outline. */
type Paint = readonly [string, string]

const RED: Paint = ['#F6888A', '#90271B']
const YELLOW: Paint = ['#FEE17E', '#D87103']
const GREEN: Paint = ['#66C07E', '#255E1B']
const GRAY: Paint = ['#A8A8A8', '#404040']
const BLACK: Paint = ['#4C4C4C', '#000000']

const ARROW = '8,1.5 14.5,8 10.5,8 10.5,14.5 5.5,14.5 5.5,8 1.5,8'
/** An up arrow rotated `deg` clockwise: 90 is right, 180 down, 45 and 135 the diagonals, drawn larger to reach the corners as Excel's do. */
const arrow = (deg: number, [fill, line]: Paint): Shape[] => {
  const turn = deg % 90 ? `translate(8 8) rotate(${deg}) scale(1.15) translate(-8 -8)` : deg ? `rotate(${deg} 8 8)` : ''
  return [['polygon', { points: ARROW, fill, stroke: line, 'stroke-linejoin': 'round', ...(turn ? { transform: turn } : {}) }]]
}
const light = ([fill, line]: Paint): Shape[] => [['circle', { cx: 8, cy: 8, r: 6.5, fill, stroke: line }]]
const dot = (fill: string): Shape[] => [['circle', { cx: 8, cy: 8, r: 6.5, fill }]]
const rimmed = (fill: string): Shape[] => [
  ['rect', { x: 1, y: 1, width: 14, height: 14, rx: 3, fill: '#1A1A1A' }],
  ['circle', { cx: 8, cy: 8, r: 4.6, fill }],
]
const flag = ([fill, line]: Paint): Shape[] => [
  ['path', { d: 'M4.5 2 13 3.6 10.8 6 13 8.4 4.8 8.6Z', fill, stroke: line, 'stroke-linejoin': 'round' }],
  ['path', { d: 'M4.2 1.5 5 15', stroke: '#000000', 'stroke-width': 1.3, 'stroke-linecap': 'round' }],
]
const CROSS = 'M5.5 5.5 10.5 10.5M10.5 5.5 5.5 10.5'
const CHECK = 'M4.8 8.3 7.2 10.6 11.3 5.6'
/** A white cross, exclamation mark or check mark on a colored disc. */
const circled = (fill: string, line: string, mark: 'x' | '!' | 'v'): Shape[] => {
  const stroke = { stroke: '#FFFFFF', 'stroke-width': 2, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', fill: 'none' }
  const sign: Shape[] =
    mark === 'x' ? [['path', { d: CROSS, ...stroke }]] : mark === 'v' ? [['path', { d: CHECK, ...stroke }]] : [['path', { d: 'M8 4.2V8.8', ...stroke }], ['circle', { cx: 8, cy: 11.6, r: 1.1, fill: '#FFFFFF' }]]
  return [['circle', { cx: 8, cy: 8, r: 6.5, fill, stroke: line }], ...sign]
}
/** A thick mark drawn twice: its outline, then its fill inside. */
const mark = (d: string, [fill, line]: Paint): Shape[] => {
  const base = { d, fill: 'none', 'stroke-linecap': 'round', 'stroke-linejoin': 'round' }
  return [['path', { ...base, stroke: line, 'stroke-width': 4.2 }], ['path', { ...base, stroke: fill, 'stroke-width': 2.4 }]]
}

/** A five-pointed star's points from the top, clockwise, alternating outer and inner. */
const STAR = Array.from({ length: 10 }, (_, k) => {
  const a = ((-90 + k * 36) * Math.PI) / 180
  const r = k % 2 ? 3.1 : 7.4
  return `${Math.round((8 + r * Math.cos(a)) * 100) / 100},${Math.round((8.6 + r * Math.sin(a)) * 100) / 100}`
})
const GOLD: Paint = ['#F8CE17', '#B8960C']
/** An empty, a half and a full star. */
const star = (filled: 0 | 1 | 2): Shape[] => {
  const all = STAR.join(' ')
  if (filled === 2) return [['polygon', { points: all, fill: GOLD[0], stroke: GOLD[1], 'stroke-linejoin': 'round' }]]
  const out: Shape[] = [['polygon', { points: all, fill: '#E6E6E6' }]]
  // the left half: the top point, then the points of the left side down to the bottom inner point
  if (filled === 1) out.push(['polygon', { points: [STAR[0], ...STAR.slice(5).reverse()].join(' '), fill: GOLD[0] }])
  out.push(['polygon', { points: all, fill: 'none', stroke: '#A6A6A6', 'stroke-linejoin': 'round' }])
  return out
}

/** Signal bars, `n` of 4 filled. */
const rating = (n: number): Shape[] =>
  [4.5, 7.5, 10.5, 13.5].map((height, i): Shape => {
    const box = { x: 1 + i * 3.7, y: 15 - height, width: 2.6, height, rx: 0.6 }
    return ['rect', i < n ? { ...box, fill: '#002E5F' } : { ...box, fill: '#D4D4D5', stroke: '#A2A2A2', 'stroke-width': 0.8 }]
  })
/** A disc filled clockwise from the top in quarters, 0 to 4. */
const quarter = (n: number): Shape[] => {
  const ring: Shape = ['circle', { cx: 8, cy: 8, r: 6.3, fill: n === 4 ? '#000000' : '#FFFFFF', stroke: '#000000', 'stroke-width': 1.4 }]
  const pie = ['', 'M8 8V1.7A6.3 6.3 0 0 1 14.3 8Z', 'M8 8V1.7A6.3 6.3 0 0 1 8 14.3Z', 'M8 8V1.7A6.3 6.3 0 1 1 1.7 8Z', ''][n]
  return pie ? [ring, ['path', { d: pie, fill: '#000000' }]] : [ring]
}
/** Four squares, `n` filled: bottom left first, then bottom right, top left, top right. */
const boxes = (n: number): Shape[] =>
  [
    [1.5, 8.5],
    [8.5, 8.5],
    [1.5, 1.5],
    [8.5, 1.5],
  ].map(([x, y], i): Shape => ['rect', { x: x!, y: y!, width: 6, height: 6, fill: i < n ? '#4F76A2' : '#E3E3E3' }])

/** Each set's icons in Excel's order: the first for the lowest numbers. */
const SETS: Record<IconSetName, Shape[][]> = {
  '3Arrows': [arrow(180, RED), arrow(90, YELLOW), arrow(0, GREEN)],
  '3ArrowsGray': [arrow(180, GRAY), arrow(90, GRAY), arrow(0, GRAY)],
  '3Flags': [flag(['#F85E5B', '#BC271F']), flag(['#FEC23F', '#BF820F']), flag(['#6BBB52', '#317C26'])],
  '3TrafficLights1': [light(RED), light(YELLOW), light(GREEN)],
  '3TrafficLights2': [rimmed('#F85E5B'), rimmed('#F4A020'), rimmed('#6BBB52')],
  '3Signs': [[['polygon', { points: '8,1 15,8 8,15 1,8', fill: '#F79A9B', stroke: '#A6332A', 'stroke-linejoin': 'round' }]], [['polygon', { points: '8,2 15,14 1,14', fill: YELLOW[0], stroke: YELLOW[1], 'stroke-linejoin': 'round' }]], light(GREEN)],
  '3Symbols': [circled('#E86A6A', RED[1], 'x'), circled('#F0A830', YELLOW[1], '!'), circled('#5BB56F', GREEN[1], 'v')],
  '3Symbols2': [mark('M4 4 12 12M12 4 4 12', ['#E07473', '#A6332A']), [...mark('M8 2.5V9', YELLOW), ['circle', { cx: 8, cy: 13, r: 1.6, fill: YELLOW[0], stroke: YELLOW[1] }]], mark('M3 8.5 6.4 11.8 13 4.2', ['#4C9956', '#255E1B'])],
  '3Stars': [star(0), star(1), star(2)],
  '3Triangles': [
    [['polygon', { points: '2,5 14,5 8,12', fill: '#E07473', stroke: '#A1463D', 'stroke-linejoin': 'round' }]],
    [['rect', { x: 2, y: 6.5, width: 12, height: 3, fill: '#F5B342', stroke: YELLOW[1] }]],
    [['polygon', { points: '8,4 14,11 2,11', fill: '#68B578', stroke: '#487C42', 'stroke-linejoin': 'round' }]],
  ],
  '4Arrows': [arrow(180, RED), arrow(135, YELLOW), arrow(45, YELLOW), arrow(0, GREEN)],
  '4ArrowsGray': [arrow(180, GRAY), arrow(135, GRAY), arrow(45, GRAY), arrow(0, GRAY)],
  '4RedToBlack': [dot('#000000'), dot('#BFBFBF'), dot('#FF8D89'), dot('#FF0000')],
  '4Rating': [rating(1), rating(2), rating(3), rating(4)],
  '4TrafficLights': [light(BLACK), light(RED), light(YELLOW), light(GREEN)],
  '5Arrows': [arrow(180, RED), arrow(135, YELLOW), arrow(90, YELLOW), arrow(45, YELLOW), arrow(0, GREEN)],
  '5ArrowsGray': [arrow(180, GRAY), arrow(135, GRAY), arrow(90, GRAY), arrow(45, GRAY), arrow(0, GRAY)],
  '5Rating': [rating(0), rating(1), rating(2), rating(3), rating(4)],
  '5Quarters': [quarter(0), quarter(1), quarter(2), quarter(3), quarter(4)],
  '5Boxes': [boxes(0), boxes(1), boxes(2), boxes(3), boxes(4)],
}

/** A conditional format's icon; nothing for a number the set does not have. */
export const CfIcon: FunctionalComponent<{ icon: Icon }> = ({ icon }) => {
  const shapes = SETS[icon.set]?.[icon.index]
  if (!shapes) return null
  return h('svg', { class: 'xs-cf-icon', viewBox: '0 0 16 16', 'aria-hidden': 'true', 'data-icon': `${icon.set}:${icon.index}` }, shapes.map(([tag, attrs]) => h(tag, attrs)))
}
CfIcon.props = ['icon']
