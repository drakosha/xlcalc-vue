// Excel's Format Painter: the selection's formats taken as a brush (with a moving border around it), then painted onto the
// next range selected with the mouse or a touch, through the source's `structure` (`copyRange` with `paste: 'formats'`, xlcalc
// 0.60), as a single undo step; once, or until Esc or the button again when it was double-clicked.
import { computed, shallowRef } from 'vue'
import { rangeAddress } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import type { FormatPainterMode } from '../contract/view'
import type { Commands } from './commands'
import type { Editing } from './editing'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

export interface PainterDeps {
  notify: (text: string) => void
}

export function usePainter(s: SheetState, sel: Selection, v: Viewport, ed: Editing, cmd: Commands, deps: PainterDeps) {
  const brush = shallowRef<{ sheet: string; rect: CellRect; mode: FormatPainterMode } | null>(null)
  /** A press on the grid while the brush was held: its release paints. */
  let armed = false

  /** The selection's formats as the brush (`off` puts it down); whether painting is on afterwards. */
  function start(mode: FormatPainterMode | 'off' = 'once'): boolean {
    const r = sel.fullRect.value
    if (mode === 'off' || !cmd.canRanges.value || !r) {
      brush.value = null
      return false
    }
    brush.value = { sheet: s.sheetName.value, rect: r, mode }
    return true
  }
  function down(): void {
    armed = !!brush.value
  }
  /** The press released: once the selection is final (a header's click selects after the release), its range is painted. */
  function up(): void {
    if (!armed) return
    armed = false
    setTimeout(() => void paint(), 0)
  }
  /** Paints the selection as a paste does: the brush tiled over a selection whose size is a multiple of it, otherwise once from its top-left cell. */
  async function paint(): Promise<void> {
    const b = brush.value
    const a = sel.fullRect.value
    if (!b || !a || !cmd.canRanges.value) return
    if (b.mode === 'once') brush.value = null
    const h = b.rect.bottom - b.rect.top + 1
    const w = b.rect.right - b.rect.left + 1
    const tiles = (a.bottom - a.top + 1) % h === 0 && (a.right - a.left + 1) % w === 0
    const target: CellRect = tiles ? a : { top: a.top, left: a.left, bottom: a.top + h - 1, right: a.left + w - 1 }
    const lock = ed.lockIn(target)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const sheet = s.sheetName.value
    await cmd.structure({ op: 'copyRange', from: { sheet: b.sheet, range: rangeAddress(b.rect) }, to: { sheet, range: rangeAddress(target) }, paste: 'formats' }, false)
    if (sheet === s.sheetName.value) sel.selectRect(target)
  }
  /** Esc puts the brush down: true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (e.key !== 'Escape' || !brush.value) return false
    e.preventDefault()
    brush.value = null
    return true
  }

  /** The moving border around the brush's range on its sheet. */
  const ants = computed(() => {
    const b = brush.value
    if (!b || b.sheet !== s.sheetName.value || s.visible.value) return null
    return v.boxOf(b.rect.top, b.rect.left, Math.min(b.rect.bottom, s.rowCount.value), Math.min(b.rect.right, s.colCount.value))
  })
  const mode = computed(() => brush.value?.mode)

  return { available: cmd.canRanges, mode, ants, start, down, up, key }
}

export type Painter = ReturnType<typeof usePainter>
