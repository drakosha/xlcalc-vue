// The selection as TSV for the clipboard; a paste is read by xlcalc's parseClipboard.

/** Rows of cells as TSV; tabs and line breaks inside a cell become spaces, as the the app copies them. */
export function toTsv(rows: string[][]): string {
  return rows.map((line) => line.map((t) => t.replace(/[\t\r\n]+/g, ' ')).join('\t')).join('\n')
}
