// The package's main entry: the sheet component, the source contract, interface texts and a source over a book in memory.
export { default as SheetView } from './components/SheetView.vue'
export type * from './source'
export type { SheetViewProps, SheetFailure } from './core/state'
export type { SheetViewState, SheetLinkEvent } from './core/sheet'
export { enMessages, ruMessages, type SheetMessages } from './messages'
export { memorySource, type MemoryBook, type MemorySheet, type MemoryCell, type MemorySource, type MemorySourceOptions } from './memory-source'
