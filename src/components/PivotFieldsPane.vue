<script setup lang="ts">
// Excel's PivotTable Fields pane over the PivotTable under the active cell: the fields with their boxes and a search, the
// Filters, Columns, Rows and Values areas with Σ Values, drag and drop between the list and the areas and within them, each
// field's menu in an area (Move Up… Remove Field, Field Settings) and in the list (Add to…), Defer Layout Update with Update.
// Each change is one `editPivotTable` through the sheet; the dialogs a menu opens are asked of the sheet (`pivotDialog`).
import { computed, nextTick, onBeforeUnmount, ref } from 'vue'
import type { PivotFieldsPaneProps } from '../contract/view'
import { check, fieldAt, fieldKind, move, remove, usedFields, type PivotArea, type PivotSpot } from '../core/pivot-layout'
import { useMessages } from '../messages'
import { IconClose, IconDown, IconFilter } from './icons'
import { usePivotPane } from './pivot-pane'
import { IconColumns, IconRows, IconSigma } from './pivot-icons'
import { useOutsidePress } from './popup'
import '../controls.css'
import '../menu.css'

const props = defineProps<PivotFieldsPaneProps>()
const m = useMessages(() => props.messages)
const pane = usePivotPane(props)
const search = ref('')

/** A field's name on the PivotTable: its custom name, or its cache field's. */
const fieldName = (i: number) => pane.definition.value?.table.fields[i]?.name ?? pane.fields.value[i]?.name ?? ''
const used = computed(() => (pane.areas.value ? usedFields(pane.areas.value) : new Set<number>()))
/** A field filtered on the PivotTable: an item hidden, or a filter by label, value or date; Excel marks it with a funnel. */
const filtered = (i: number) => !!pane.definition.value?.table.fields[i]?.items?.some((x) => x.hidden) || !!pane.definition.value?.table.filters?.some((f) => f.field === i)
const list = computed(() => {
  const q = search.value.trim().toLocaleLowerCase()
  return pane.fields.value.map((f, i) => ({ i, name: f.name })).filter((f) => !q || f.name.toLocaleLowerCase().includes(q))
})

const AREAS: { area: PivotArea; label: () => string; icon: typeof IconSigma }[] = [
  { area: 'filters', label: () => m.value.areaFilters, icon: IconFilter },
  { area: 'columns', label: () => m.value.areaColumns, icon: IconColumns },
  { area: 'rows', label: () => m.value.areaRows, icon: IconRows },
  { area: 'values', label: () => m.value.areaValues, icon: IconSigma },
]
/** An area's chips: the label of each field in it. */
function chips(area: PivotArea): string[] {
  const a = pane.areas.value
  if (!a) return []
  if (area === 'filters') return a.filters.map((p) => fieldName(p.field))
  if (area === 'values') return a.values.map((v) => v.name ?? fieldName(v.field))
  return a[area].map((f) => (f === 'values' ? m.value.fieldListValues : fieldName(f)))
}

// ── menus: a field's in an area, and a field's in the list ──
interface MenuItem {
  id: string
  label: string
  run: () => void
  disabled?: boolean
  group?: boolean
}
const menu = ref<{ items: MenuItem[]; left: number; top: number; back: HTMLElement | null } | null>(null)
const menuEl = ref<HTMLElement | null>(null)
useOutsidePress(menuEl, () => (menu.value = null), () => !!menu.value)

function openMenu(items: MenuItem[], from: HTMLElement): void {
  const r = from.getBoundingClientRect()
  menu.value = { items, left: r.left, top: r.bottom + 2, back: from }
  void nextTick(() => {
    const el = menuEl.value
    if (!el) return
    // kept in the viewport: above the control where there is no room below
    const box = el.getBoundingClientRect()
    if (box.bottom > window.innerHeight && r.top - box.height - 2 >= 0) menu.value = menu.value && { ...menu.value, top: r.top - box.height - 2 }
    if (box.right > window.innerWidth) menu.value = menu.value && { ...menu.value, left: Math.max(0, window.innerWidth - box.width - 4) }
    el.querySelector<HTMLButtonElement>('[role="menuitem"]:not(:disabled)')?.focus()
  })
}
function closeMenu(): void {
  const back = menu.value?.back
  menu.value = null
  back?.focus()
}
function choose(item: MenuItem): void {
  if (item.disabled) return
  closeMenu()
  item.run()
}
/** Up and Down move among the menu's items, Home and End to the ends, Esc and Tab close it. */
function onMenuKey(e: KeyboardEvent): void {
  const items = [...(menuEl.value?.querySelectorAll<HTMLButtonElement>('[role="menuitem"]:not(:disabled)') ?? [])]
  const at = items.indexOf(document.activeElement as HTMLButtonElement)
  const go = (i: number) => items[(i + items.length) % items.length]?.focus()
  if (e.key === 'ArrowDown') go(at + 1)
  else if (e.key === 'ArrowUp') go(at - 1)
  else if (e.key === 'Home') go(0)
  else if (e.key === 'End') go(items.length - 1)
  else if (e.key === 'Escape' || e.key === 'Tab') closeMenu()
  else return
  e.preventDefault()
}

const apply = (next: ReturnType<typeof move>) => pane.apply(next)
const moveTo = (from: PivotSpot, to: PivotArea, index?: number) => pane.areas.value && apply(move(pane.areas.value, from, to, index, pane.fields.value))

/** Excel's menu of a field in an area: its moves within the area, to the other areas, Remove Field, and its settings. */
function chipItems(area: PivotArea, index: number, at: DOMRect): MenuItem[] {
  const a = pane.areas.value!
  const n = chips(area).length
  const spot: PivotSpot = { area, index }
  const field = fieldAt(a, spot)
  const sigma = field === 'values'
  const x = m.value
  const settings: MenuItem =
    area === 'values'
      ? { id: 'valueFieldSettings', label: x.valueFieldSettings, group: true, run: () => pane.dialog({ dialog: 'valueFieldSettings', field: a.values[index]!.field, value: index, at: rectOf(at) }) }
      : { id: 'fieldSettings', label: x.fieldSettings, group: true, disabled: sigma, run: () => typeof field === 'number' && pane.dialog({ dialog: 'fieldSettings', field, at: rectOf(at) }) }
  return [
    { id: 'moveUp', label: x.fieldMoveUp, disabled: index === 0, run: () => moveTo(spot, area, index - 1) },
    { id: 'moveDown', label: x.fieldMoveDown, disabled: index === n - 1, run: () => moveTo(spot, area, index + 2) },
    { id: 'moveFirst', label: x.fieldMoveFirst, disabled: index === 0, run: () => moveTo(spot, area, 0) },
    { id: 'moveLast', label: x.fieldMoveLast, disabled: index === n - 1, run: () => moveTo(spot, area) },
    { id: 'toFilters', label: x.fieldToFilters, group: true, disabled: area === 'filters' || sigma, run: () => moveTo(spot, 'filters') },
    { id: 'toRows', label: x.fieldToRows, disabled: area === 'rows', run: () => moveTo(spot, 'rows') },
    { id: 'toColumns', label: x.fieldToColumns, disabled: area === 'columns', run: () => moveTo(spot, 'columns') },
    { id: 'toValues', label: x.fieldToValues, disabled: area === 'values' || sigma, run: () => moveTo(spot, 'values') },
    { id: 'remove', label: x.fieldRemove, group: true, run: () => apply(remove(a, spot)) },
    settings,
  ]
}
/** Excel's menu of a field in the list: Add to each area, disabled for the axis it is on already. */
function listItems(field: number): MenuItem[] {
  const a = pane.areas.value!
  const spot: PivotSpot = { area: 'list', field }
  const x = m.value
  return [
    { id: 'addFilters', label: x.fieldAddFilters, disabled: a.filters.some((p) => p.field === field), run: () => moveTo(spot, 'filters') },
    { id: 'addRows', label: x.fieldAddRows, disabled: a.rows.includes(field), run: () => moveTo(spot, 'rows') },
    { id: 'addColumns', label: x.fieldAddColumns, disabled: a.columns.includes(field), run: () => moveTo(spot, 'columns') },
    { id: 'addValues', label: x.fieldAddValues, run: () => moveTo(spot, 'values') },
  ]
}
const rectOf = (r: DOMRect) => ({ left: r.left, top: r.top, width: r.width, height: r.height })

function onChipClick(e: MouseEvent, area: PivotArea, index: number): void {
  if (dragged || !pane.editable.value) return
  const el = e.currentTarget as HTMLElement
  openMenu(chipItems(area, index, el.getBoundingClientRect()), el)
}
function onListMenu(e: Event, field: number): void {
  if (!pane.editable.value || !pane.areas.value) return
  e.preventDefault()
  const el = (e.currentTarget as HTMLElement).querySelector<HTMLElement>('input') ?? (e.currentTarget as HTMLElement)
  openMenu(listItems(field), el)
}
function onListKey(e: KeyboardEvent, field: number): void {
  if (e.key === 'ContextMenu' || (e.key === 'F10' && e.shiftKey)) onListMenu(e, field)
}
function onCheck(field: number, on: boolean): void {
  if (pane.areas.value) apply(check(pane.areas.value, field, on, pane.fields.value))
}
function onFilter(e: MouseEvent, field: number): void {
  pane.dialog({ dialog: 'fieldFilter', field, at: rectOf((e.currentTarget as HTMLElement).getBoundingClientRect()) })
}

// ── drag and drop: a field of the list or a chip, onto an area at a place, or a chip out of the areas to remove it ──
const DRAG_PX = 4
const root = ref<HTMLElement | null>(null)
const drag = ref<{ from: PivotSpot; label: string; x: number; y: number; over: { area: PivotArea; index: number } | null } | null>(null)
let press: { from: PivotSpot; label: string; x: number; y: number; id: number } | null = null
/** A drag just ended: the click it ends with opens no menu. */
let dragged = false

function onDown(e: PointerEvent, from: PivotSpot, label: string): void {
  if (e.button !== 0 || !pane.editable.value) return
  press = { from, label, x: e.clientX, y: e.clientY, id: e.pointerId }
  dragged = false
  window.addEventListener('pointermove', onMove)
  window.addEventListener('pointerup', onUp)
  window.addEventListener('pointercancel', endDrag)
}
/** The area under the point and the place among its chips: before the first chip whose middle is below it, or at the end. */
function dropAt(x: number, y: number): { area: PivotArea; index: number } | null {
  const el = document.elementFromPoint(x, y)?.closest<HTMLElement>('[data-xs-pf-area]')
  if (!el || !root.value?.contains(el)) return null
  const area = el.dataset.xsPfArea as PivotArea
  const items = [...el.querySelectorAll<HTMLElement>('[data-xs-pf-chip]')]
  const index = items.findIndex((c) => {
    const r = c.getBoundingClientRect()
    return r.top + r.height / 2 > y
  })
  return { area, index: index < 0 ? items.length : index }
}
function onMove(e: PointerEvent): void {
  if (!press || e.pointerId !== press.id) return
  if (!drag.value && Math.hypot(e.clientX - press.x, e.clientY - press.y) < DRAG_PX) return
  e.preventDefault()
  drag.value = { from: press.from, label: press.label, x: e.clientX, y: e.clientY, over: dropAt(e.clientX, e.clientY) }
}
function onUp(e: PointerEvent): void {
  const d = drag.value
  endDrag()
  if (!d) return
  dragged = true
  setTimeout(() => (dragged = false))
  const over = dropAt(e.clientX, e.clientY)
  if (over) moveTo(d.from, over.area, over.index)
  else if (d.from.area !== 'list' && pane.areas.value) apply(remove(pane.areas.value, d.from))
}
function endDrag(): void {
  press = null
  drag.value = null
  window.removeEventListener('pointermove', onMove)
  window.removeEventListener('pointerup', onUp)
  window.removeEventListener('pointercancel', endDrag)
}
function onKey(e: KeyboardEvent): void {
  if (e.key === 'Escape' && drag.value) {
    e.preventDefault()
    endDrag()
  }
}
onBeforeUnmount(endDrag)

const dropBefore = (area: PivotArea, index: number) => drag.value?.over?.area === area && drag.value.over.index === index
const kindOf = (i: number) => {
  const f = pane.fields.value[i]
  return f ? fieldKind(f) : 'text'
}
</script>

<template>
  <aside ref="root" class="xs-pf" :class="{ 'xs-pf-dragging': !!drag }" :aria-label="m.fieldListTitle" data-xs-pf @keydown="onKey">
    <div class="xs-pf-head">
      <span class="xs-pf-title">{{ m.fieldListTitle }}</span>
      <button type="button" class="xs-tool" :title="m.close" :aria-label="m.close" data-xs-pf-close @click="sheet?.fieldList(false)">
        <IconClose />
      </button>
    </div>
    <p class="xs-pf-note">{{ m.fieldListChoose }}</p>
    <input v-model="search" class="xs-field xs-pf-search" type="search" :placeholder="m.fieldListSearch" :aria-label="m.fieldListSearch" autocomplete="off" spellcheck="false" data-xs-pf-search />
    <ul class="xs-pf-fields" role="list" data-xs-pf-fields>
      <li
        v-for="f in list"
        :key="f.i"
        class="xs-pf-field"
        :data-xs-pf-field="f.i"
        :data-kind="kindOf(f.i)"
        @pointerdown="onDown($event, { area: 'list', field: f.i }, f.name)"
        @contextmenu="onListMenu($event, f.i)"
      >
        <label class="xs-pf-check">
          <input type="checkbox" :checked="used.has(f.i)" :disabled="!pane.editable.value" @change="onCheck(f.i, ($event.target as HTMLInputElement).checked)" @keydown="onListKey($event, f.i)" />
          <span class="xs-pf-name" :class="{ 'xs-pf-on': used.has(f.i) }">{{ f.name }}</span>
        </label>
        <IconFilter v-if="filtered(f.i)" class="xs-pf-funnel" data-xs-pf-filtered />
        <button type="button" class="xs-tool xs-pf-arrow" :aria-label="`${m.fieldListFilter}: ${f.name}`" :disabled="!pane.editable.value" data-xs-pf-filter @pointerdown.stop @click="onFilter($event, f.i)">
          <IconDown />
        </button>
      </li>
    </ul>
    <p class="xs-pf-note">{{ m.fieldListDrag }}</p>
    <div class="xs-pf-areas">
      <section v-for="x in AREAS" :key="x.area" class="xs-pf-area" :class="{ 'xs-pf-drop-end': dropBefore(x.area, chips(x.area).length) }" :data-xs-pf-area="x.area" :aria-label="x.label()">
        <div class="xs-pf-area-head">
          <component :is="x.icon" />
          <span>{{ x.label() }}</span>
        </div>
        <ul class="xs-pf-chips" role="list">
          <li v-for="(label, i) in chips(x.area)" :key="`${i}:${label}`" :class="{ 'xs-pf-drop-before': dropBefore(x.area, i) }" data-xs-pf-chip>
            <button
              type="button"
              class="xs-pf-chip"
              :aria-haspopup="pane.editable.value ? 'menu' : undefined"
              :data-xs-pf-chip-of="x.area"
              @pointerdown="onDown($event, { area: x.area, index: i }, label)"
              @click="onChipClick($event, x.area, i)"
            >
              <span class="xs-pf-chip-label">{{ label }}</span>
              <IconDown />
            </button>
          </li>
        </ul>
      </section>
    </div>
    <div class="xs-pf-foot">
      <label class="xs-pf-check">
        <input type="checkbox" :checked="pane.defer.value" :disabled="!pane.editable.value" data-xs-pf-defer @change="pane.setDefer(($event.target as HTMLInputElement).checked)" />
        <span>{{ m.fieldListDefer }}</span>
      </label>
      <button type="button" class="xs-button" :disabled="!pane.waiting.value" data-xs-pf-update @click="pane.update()">{{ m.fieldListUpdate }}</button>
    </div>
    <div v-if="menu" ref="menuEl" class="xs-menu xs-pf-menu" role="menu" :style="{ left: `${menu.left}px`, top: `${menu.top}px` }" data-xs-pf-menu @keydown="onMenuKey">
      <template v-for="item in menu.items" :key="item.id">
        <div v-if="item.group" class="xs-menu-sep" role="separator" />
        <button type="button" class="xs-menu-item" role="menuitem" :disabled="item.disabled" :data-xs-pf-item="item.id" @click="choose(item)">{{ item.label }}</button>
      </template>
    </div>
    <div v-if="drag" class="xs-pf-ghost" :style="{ left: `${drag.x + 8}px`, top: `${drag.y + 8}px` }" aria-hidden="true">{{ drag.label }}</div>
  </aside>
</template>

<style src="../pivot-fields.css"></style>
