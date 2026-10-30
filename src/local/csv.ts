// A CSV file into a workbook of one sheet and a sheet into a CSV file, through `xlcalc/csv`, which loads only when used.
import type { Locale } from 'xlcalc'
import type { CsvRead, CsvWritten, ReadCsvOptions, WriteCsvOptions } from 'xlcalc/csv'
import { loadLocale } from './locale'
import type { LocalSource } from './types'

/** A file's bytes or text: a `File` from a file input or a drop, the bytes themselves, or text already decoded. */
type CsvData = Blob | ArrayBuffer | Uint8Array | string

/** A tag's locale module loaded first: the library accepts a tag it has loaded and rejects others. */
async function loaded(locale: string | Locale | undefined): Promise<void> {
  if (typeof locale === 'string') await loadLocale(locale)
}

/**
 * Reads a CSV file into a workbook of one sheet as Excel opens it in `locale` (xlcalc's `readCsv`): each field as text typed
 * into a cell there, the encoding and the delimiter found. A `File` names the sheet after itself unless `name` is given.
 */
export async function openCsv(data: CsvData, opts: ReadCsvOptions = {}): Promise<CsvRead> {
  const named = opts.name === undefined && typeof File !== 'undefined' && data instanceof File ? { name: data.name } : {}
  const [bytes, { readCsv }] = await Promise.all([data instanceof Blob ? data.arrayBuffer() : data, import('xlcalc/csv'), loaded(opts.locale)])
  return readCsv(bytes, { ...opts, ...named })
}

/**
 * Writes a sheet of the source's workbook (`sheet`, the workbook's active one by default) as Excel saves it as CSV (xlcalc's
 * `writeCsv`), its formulas by their values. In Manual the source calculates first, as Excel recalculates before saving.
 */
export async function saveCsv(source: LocalSource, opts: WriteCsvOptions = {}): Promise<CsvWritten> {
  await source.calculate()
  const [snapshot, { writeCsv }] = await Promise.all([source.snapshot(), import('xlcalc/csv'), loaded(opts.locale)])
  return writeCsv(snapshot, opts)
}
