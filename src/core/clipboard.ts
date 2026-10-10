// Clipboard text: the selection as TSV for Excel, and TSV from Excel (quoted fields included) as a rectangle.

/** Rows of cells as TSV; tabs and line breaks inside a cell become spaces, as the the app copies them. */
export function toTsv(rows: string[][]): string {
  return rows.map((line) => line.map((t) => t.replace(/[\t\r\n]+/g, ' ')).join('\t')).join('\n')
}

/** TSV as Excel puts it on the clipboard: a quoted field may hold tabs, line breaks and doubled quotes; the last line break ends the data. */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let i = 0
  const src = text.replace(/\r\n?/g, '\n')
  while (i < src.length) {
    const ch = src[i]!
    if (ch === '"' && field === '') {
      // A quoted field: up to the closing quote, a doubled quote is one quote.
      let j = i + 1
      let quoted = ''
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            quoted += '"'
            j += 2
            continue
          }
          break
        }
        quoted += src[j]
        j++
      }
      const next = src[j + 1]
      if (j < src.length && (next === undefined || next === '\t' || next === '\n')) {
        field = quoted
        i = j + 1
        continue
      }
    }
    if (ch === '\t') {
      row.push(field)
      field = ''
    } else if (ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else field += ch
    i++
  }
  if (field !== '' || row.length) {
    row.push(field)
    rows.push(row)
  }
  return rows
}
