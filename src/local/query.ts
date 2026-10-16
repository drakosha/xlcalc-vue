// A sheet's shown texts and values as the display filter reads them (`../sheet-query`): each column read from the
// session when a condition, a sort or a funnel first needs it, kept in arrays by row and brought up to date by edits.
import type { CellValue, CellView } from 'xlcalc'
import { colLetters, parseCell } from 'xlcalc/a1'
import { throwIfAborted } from '../abort'
import type { QuerySheet } from '../sheet-query'

/** What a cell holds, for its value. */
const EMPTY = 0
const NUMBER = 1
const TEXT = 2
const BOOLEAN = 3
const ERROR = 4

const isEmpty = (value: CellValue | undefined): boolean => value === null || value === undefined || value === ''

/** A column's cells by row: the shown texts, their values' kinds, numbers (a boolean as 0 or 1), and a text value its shown text does not say or an error's code. */
class Column {
  texts: string[]
  kinds: Uint8Array
  nums: Float64Array
  readonly other = new Map<number, string>()

  constructor(size: number) {
    this.texts = new Array<string>(size).fill('')
    this.kinds = new Uint8Array(size)
    this.nums = new Float64Array(size)
  }

  set(row: number, text: string, value: CellValue | undefined): void {
    if (row >= this.kinds.length) this.grow(row + 1)
    this.texts[row] = text
    this.other.delete(row)
    if (isEmpty(value)) this.kinds[row] = EMPTY
    else if (typeof value === 'number') {
      this.kinds[row] = NUMBER
      this.nums[row] = value
    } else if (typeof value === 'boolean') {
      this.kinds[row] = BOOLEAN
      this.nums[row] = value ? 1 : 0
    } else if (typeof value === 'string') {
      this.kinds[row] = TEXT
      if (value !== text) this.other.set(row, value)
    } else {
      this.kinds[row] = ERROR
      this.other.set(row, value!.error)
    }
  }

  valueAt(row: number): CellValue {
    switch (this.kinds[row]) {
      case NUMBER:
        return this.nums[row]!
      case BOOLEAN:
        return this.nums[row] === 1
      case TEXT:
        return this.other.get(row) ?? this.texts[row]!
      case ERROR:
        return { error: this.other.get(row) } as CellValue
      default:
        return null
    }
  }

  /** Room for `size` rows at least, doubling, so that edits past the end do not copy at each row. */
  private grow(size: number): void {
    const n = Math.max(size, this.kinds.length * 2)
    const kinds = new Uint8Array(n)
    kinds.set(this.kinds)
    const nums = new Float64Array(n)
    nums.set(this.nums)
    for (let i = this.texts.length; i < n; i++) this.texts.push('')
    this.kinds = kinds
    this.nums = nums
  }
}

/** A column being read: when it is done, the signal that stops the read, and the changes that came meanwhile. */
interface Reading {
  done: Promise<void>
  signal?: AbortSignal
  missed: [row: number, text: string, value: CellValue | undefined][]
}

/** A sheet's columns as the display filter has needed them, read through `read` (the session's `getCells` over ranges, stopped by a signal). */
export class SheetValues implements QuerySheet {
  /** Columns read, and those being read. */
  private readonly cols = new Map<number, Column>()
  private readonly reading = new Map<number, Reading>()

  /** `maxRow`: where the sheet's cells end; an edit past it moves it on, as Excel's used range grows. */
  constructor(
    public maxRow: number,
    private readonly read: (ranges: string[], signal?: AbortSignal) => Promise<Record<string, CellView>>,
  ) {}

  /**
   * The columns read, the missing ones in one call that `signal` stops. A read another request started and its signal
   * stopped is made again for this one.
   */
  async load(cols: number[], signal?: AbortSignal): Promise<void> {
    for (;;) {
      const want = cols.filter((c) => !this.cols.has(c) && !this.reading.has(c))
      if (want.length) this.start(want, signal)
      const pending = cols.flatMap((c) => this.reading.get(c) ?? [])
      try {
        await Promise.all(pending.map((r) => r.done))
        return
      } catch (e) {
        // the session refuses a stopped call with an error of its own; the component expects an abort
        throwIfAborted(signal)
        if (!pending.some((r) => r.signal?.aborted)) throw e
      }
    }
  }

  /** Reads columns `want` in one call. */
  private start(want: number[], signal: AbortSignal | undefined): void {
    const rows = this.maxRow
    const views = rows ? this.read(want.map((c) => `${colLetters(c)}1:${colLetters(c)}${rows}`), signal) : Promise.resolve<Record<string, CellView>>({})
    const done = views.then(
      (cells) => {
        const made = new Map(want.map((c) => [c, new Column(rows + 1)]))
        for (const [addr, v] of Object.entries(cells)) {
          const at = parseCell(addr)
          if (at) made.get(at.col)?.set(at.row, v.text ?? '', v.value)
        }
        for (const [c, column] of made) {
          for (const [row, text, value] of this.reading.get(c)!.missed) column.set(row, text, value)
          this.cols.set(c, column)
          this.reading.delete(c)
        }
      },
      (e: unknown) => {
        for (const c of want) this.reading.delete(c)
        throw e
      },
    )
    for (const c of want) this.reading.set(c, { done, signal, missed: [] })
  }

  /** A cell's new text and value, kept where its column is read or being read. */
  set(row: number, col: number, text: string, value: CellValue | undefined): void {
    if (row > this.maxRow && (text !== '' || !isEmpty(value))) this.maxRow = row
    const column = this.cols.get(col)
    if (column) column.set(row, text, value)
    else this.reading.get(col)?.missed.push([row, text, value])
  }

  textAt(row: number, col: number): string {
    return this.cols.get(col)?.texts[row] ?? ''
  }

  valueAt(row: number, col: number): CellValue | undefined {
    return this.cols.get(col)?.valueAt(row) ?? null
  }
}
