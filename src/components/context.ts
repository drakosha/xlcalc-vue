// The sheet's context, shared by the view's parts (formula bar, grid, tabs).
import { inject, type InjectionKey } from 'vue'
import type { SheetContext } from '../core/sheet'

export const sheetContextKey: InjectionKey<SheetContext> = Symbol('xs-sheet')

export function useSheetContext(): SheetContext {
  const ctx = inject(sheetContextKey)
  if (!ctx) throw new Error('SheetView parts work only inside SheetView')
  return ctx
}
