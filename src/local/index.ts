// The `xlcalc-sheet/local` entry: a sheet source over an xlcalc book computed in the browser, and xlsx in and out.
import type { BookInput } from 'xlcalc'
import type { XlsxIssue } from 'xlcalc/xlsx'
import type { LocalSource, LocalSourceOptions } from './types'

export type * from './types'

export function localSource(_book: BookInput, _opts: LocalSourceOptions = {}): LocalSource {
  throw new Error('not implemented')
}

export async function openXlsx(_data: Blob | ArrayBuffer | Uint8Array): Promise<{ book: BookInput; issues: XlsxIssue[] }> {
  throw new Error('not implemented')
}

export async function saveXlsx(_source: LocalSource): Promise<{ data: Uint8Array; issues: XlsxIssue[] }> {
  throw new Error('not implemented')
}

export function groupIssues(_issues: XlsxIssue[]): { code: string; count: number; sheets: string[]; message: string }[] {
  throw new Error('not implemented')
}
