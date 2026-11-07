// The display filter's supported subset of xlcalc's stored criteria. Other criteria stay opaque and are never approximated.
import type { ColumnFilter, FilterCriteria } from 'xlcalc/contract'
import { parseRange, rangeAddress } from 'xlcalc/a1'
import type { SheetFilter } from './contract/source'

type Condition = NonNullable<SheetFilter['cols']>[number]

/** An Excel wildcard pattern that is exactly a literal substring, with escaped ~, * and ?. */
function substring(pattern: string): string | null {
  if (!pattern.startsWith('*') || !pattern.endsWith('*')) return null
  const body = pattern.slice(1, -1)
  let text = ''
  for (let i = 0; i < body.length; i++) {
    const c = body[i]!
    if (c === '~') {
      const next = body[++i]
      if (!next || !'~*?'.includes(next)) return null
      text += next
    } else if (c === '*' || c === '?') return null
    else text += c
  }
  // The display filter trims its input; preserve patterns whose spaces would change their meaning as opaque.
  return text && text.trim() === text ? text : null
}

export function displayCondition(filter: ColumnFilter): Condition | null {
  if (filter.type === 'values' && !filter.dates?.length) return { values: [...(filter.values ?? []), ...(filter.blank ? [''] : [])] }
  if (filter.type === 'custom' && filter.conditions.length === 1) {
    const test = filter.conditions[0]!
    const text = !test.operator || test.operator === 'equal' ? substring(test.value) : null
    if (text !== null) return { text }
  }
  return null
}

export const unsupportedColumns = (criteria: FilterCriteria | undefined): number[] => (criteria?.columns ?? []).filter(c => c.filter && !displayCondition(c.filter)).map(c => c.column)

/** The supported conditions, addressed by sheet column; complex sorts stay in the original criteria. */
export function displayFilter(range: string, criteria: FilterCriteria | undefined): SheetFilter {
  const r = parseRange(range)!
  const cols: NonNullable<SheetFilter['cols']> = {}
  for (const c of criteria?.columns ?? []) {
    const condition = c.filter && displayCondition(c.filter)
    if (condition) cols[r.left + c.column] = condition
  }
  const s = criteria?.sort
  const by = s?.conditions?.length === 1 ? s.conditions[0] : undefined
  const column = by && parseRange(by.range)
  const sort = s && by && column && column.left === column.right && !s.byColumns && !s.caseSensitive && !s.method && !by.customList && (!by.by || by.by === 'value') ? { col: column.left, dir: by.descending ? 'desc' as const : 'asc' as const } : undefined
  return { head: r.top, last: r.bottom, ...(Object.keys(cols).length ? { cols } : {}), ...(sort ? { sort } : {}) }
}

export function withColumn(criteria: FilterCriteria | undefined, column: number, condition: Condition | undefined): FilterCriteria {
  const columns = (criteria?.columns ?? []).filter(c => c.column !== column)
  const original = criteria?.columns?.find(c => c.column === column)
  const filter: ColumnFilter | undefined = condition?.values ? { type: 'values', values: condition.values.filter(v => v !== '' && (!condition.text || v.toLocaleLowerCase().includes(condition.text.toLocaleLowerCase()))), ...(condition.values.includes('') && !condition.text ? { blank: true } : {}) }
    : condition?.text ? { type: 'custom', conditions: [{ operator: 'equal', value: `*${condition.text.replace(/[~*?]/g, '~$&')}*` }] } : undefined
  const { filter: _old, ...button } = original ?? { column }
  if (filter || original?.hiddenButton !== undefined || original?.showButton !== undefined) columns.push({ ...button, ...(filter ? { filter } : {}) })
  columns.sort((a, b) => a.column - b.column)
  return { ...criteria, columns }
}

export function withSort(criteria: FilterCriteria | undefined, range: string, sort: SheetFilter['sort'] | null): FilterCriteria {
  const { sort: _old, ...rest } = criteria ?? {}
  if (!sort) return rest
  const r = parseRange(range)!
  const data = { ...r, top: Math.min(r.top + 1, r.bottom) }
  return { ...rest, sort: { range: rangeAddress(data), conditions: [{ range: rangeAddress({ ...data, left: sort.col, right: sort.col }), ...(sort.dir === 'desc' ? { descending: true } : {}) }] } }
}
