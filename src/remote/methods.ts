// The source's methods as the transport carries them, shared by `./remote` and `./server`: which change the workbook,
// which return an `EditOutcome`, and where a method takes its `AbortSignal`, which does not cross the network.
import type { SheetSource, SourceEvent } from '../contract/source'

/** The methods a client calls; `subscribe` is the event stream, and the readonly fields come with the description. */
export type CallName = Exclude<{ [K in keyof SheetSource]-?: NonNullable<SheetSource[K]> extends (...args: never[]) => unknown ? K : never }[keyof SheetSource], 'subscribe'>

export interface CallKind {
  /** It changes the workbook: allowed with write access only. */
  write?: true
  /** It returns an `EditOutcome`: a failure comes back as one rather than as an error. */
  outcome?: true
  /** The argument that is the signal, or with `inOpts` the options object that carries it. */
  signal?: number
  inOpts?: true
}

/** Every method of `SheetSource` but `subscribe`: a method the contract adds fails the type check until it is listed here. */
export const CALLS: Readonly<Record<CallName, CallKind>> = {
  sheets: { signal: 0 },
  window: { signal: 1 },
  edit: { write: true, outcome: true },
  paste: { write: true, outcome: true },
  fill: { write: true, outcome: true },
  undo: { write: true, outcome: true },
  redo: { write: true, outcome: true },
  structure: { write: true, outcome: true },
  layout: { write: true, outcome: true },
  copyText: { signal: 2, inOpts: true },
  format: { write: true, outcome: true },
  validation: { signal: 2 },
  pick: { write: true, outcome: true },
  freeze: { write: true, outcome: true },
  tabColor: { write: true, outcome: true },
  show: { write: true, outcome: true },
  functionNames: { signal: 0 },
  functionSignature: { signal: 1 },
  explain: {},
  end: { signal: 3 },
  stats: { signal: 3 },
  find: { signal: 1 },
  replace: { write: true, outcome: true },
  outline: { write: true, outcome: true },
  visibleRows: { signal: 2 },
  columnValues: { signal: 3 },
  setAutoFilter: { write: true, outcome: true },
  clear: { write: true, outcome: true },
  link: { write: true, outcome: true },
  note: { write: true, outcome: true },
  names: { signal: 0 },
  setNames: { write: true, outcome: true },
  dependents: { signal: 2 },
  rules: { signal: 1 },
  setRules: { write: true, outcome: true },
  table: { write: true, outcome: true },
  spelling: { signal: 1 },
  learnWord: { write: true },
  autoSum: { signal: 2 },
  calculation: { signal: 0 },
  setCalculation: { write: true, outcome: true },
  calculate: { write: true, outcome: true },
  protect: { write: true, outcome: true },
  specialCells: { signal: 3 },
  namesFromSelection: { signal: 3 },
}

export const isCall = (name: string): name is CallName => Object.hasOwn(CALLS, name)

/** What the server tells a client about its source: the methods the client may call and the source's readonly fields. */
export interface Manifest {
  protocol: 1
  access: 'read' | 'write'
  /** The callable methods, and `subscribe` when the source has events. */
  methods: string[]
  maxCells?: number
  autoSumFunctions?: SheetSource['autoSumFunctions']
  history?: { undo: number; redo: number }
  locale?: string
  /** With events: this server's mark and the number of the source's events so far, to tell whether a stream missed some. */
  epoch?: string
  seq?: number
}

/** A call's body: the arguments, with the indexes of those that were `undefined`, which JSON would turn into null. */
export interface CallBody {
  method: string
  args: unknown[]
  undef?: number[]
}

/** A call's response: the result (none for `undefined`), and the history after it for a source that has one; or the error's message. */
export type CallResponse = { result?: unknown; history?: { undo: number; redo: number } } | { error: string }

/** The most calls one batch takes: a client sends more as several. */
export const BATCH_MAX = 1000

/** A line of a batch's response: the call's index in the batch, its status and its response. */
export type BatchLine = { i: number; status: number } & CallResponse

/** A message of the stream: the source's event with the history after it, or without one the count the stream starts from. */
export interface StreamEvent {
  event?: SourceEvent
  history?: { undo: number; redo: number }
  epoch: string
  seq: number
}

/** The history as plain numbers, read from a live getter. */
export const historyOf = (source: Pick<SheetSource, 'history'>): Manifest['history'] => (source.history ? { undo: source.history.undo, redo: source.history.redo } : undefined)
