// Production app, its real workers and both viewport sizes: rich fonts and stored filters through open/edit/save.
import assert from 'node:assert/strict'
import { mkdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { chromium } from 'playwright'
import { preview } from 'vite'
import { compute, type BookInput, type FilterCriteria } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const out = resolve('test-results/screens')
mkdirSync(out, { recursive: true })
const criteria: FilterCriteria = { columns: [{ column: 1, filter: { type: 'values', values: ['a'] } }] }
const runs = [{ start: 0, font: { bold: true } }, { start: 2, font: { italic: true, color: '#FF0000' } }]
const book = (saved = criteria): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 'Text' }, B1: { value: 'Status' }, A2: { value: '😀Hello', richText: runs }, B2: { value: 'a' }, A3: { value: 'other' }, B3: { value: 'b' } }, columns: { A: { width: 20 } }, rows: { '3': { hidden: true } }, filtered: true, autoFilter: { range: 'A1:B3', criteria: saved } }] })
async function file(saved = criteria) {
  const input = book(saved)
  return { name: 'rich-filter.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from((await writeXlsx({ book: input, result: compute(input) })).data) }
}
const server = await preview({ configFile: resolve('vite.example.config.ts'), preview: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
const browser = await chromium.launch()
try {
  for (const size of [{ name: 'desktop', viewport: { width: 1920, height: 1200 }, hasTouch: false }, { name: 'phone', viewport: { width: 390, height: 844 }, hasTouch: true }]) {
    const page = await browser.newPage(size)
    const errors: string[] = []
    page.on('pageerror', e => errors.push(e.message))
    const sourceBook = () => page.evaluate(() => (window as unknown as { __spa: { book(): Promise<BookInput> } }).__spa.book())
    await page.goto(`${server.resolvedUrls!.local[0]}spa.html`)
    await page.locator('[data-spa-file]').setInputFiles(await file())
    await page.waitForFunction(() => document.querySelectorAll('.xs-rich-run').length === 2)
    const fonts = await page.locator('.xs-rich-run').evaluateAll(elements => elements.map(el => { const s = getComputedStyle(el); return { text: el.textContent, weight: s.fontWeight, italic: s.fontStyle, color: s.color } }))
    assert.deepEqual(fonts, [{ text: '😀', weight: '700', italic: 'normal', color: 'rgb(31, 35, 40)' }, { text: 'Hello', weight: '400', italic: 'italic', color: 'rgb(255, 0, 0)' }])
    await page.screenshot({ path: resolve(out, `rich-filter-${size.name}.png`) })
    await page.locator('[data-xs-af="2"]').click()
    await page.locator('.xs-value').filter({ has: page.getByText('a', { exact: true }) }).locator('input').uncheck()
    await page.locator('.xs-value').filter({ has: page.getByText('b', { exact: true }) }).locator('input').check()
    await page.locator('[data-xs-apply]').click()
    await page.waitForFunction(async () => { const b = await (window as unknown as { __spa: { book(): Promise<BookInput> } }).__spa.book(); return b.sheets[0]!.rows?.['2']?.hidden && !b.sheets[0]!.rows?.['3']?.hidden })
    const [download] = await Promise.all([page.waitForEvent('download'), page.locator('[data-spa="save"]').click()])
    const saved = resolve(out, `rich-filter-${size.name}.xlsx`)
    await download.saveAs(saved)
    const opened = await readXlsx(new Uint8Array(readFileSync(saved)))
    assert.deepEqual(opened.book.sheets[0]!.autoFilter?.criteria?.columns?.[0]?.filter, { type: 'values', values: ['b'] })
    assert.equal(opened.book.sheets[0]!.cells.A2!.richText?.[0]?.font?.bold, true)
    await page.locator('[data-xs-filter-clear]').click()
    await page.waitForFunction(() => document.querySelectorAll('.xs-rich-run').length === 2)
    assert.equal((await sourceBook()).sheets[0]!.rows?.['2']?.hidden ?? false, false)
    const opaque: FilterCriteria = { columns: [{ column: 1, filter: { type: 'top', value: 10 } }] }
    await page.locator('[data-spa-file]').setInputFiles(await file(opaque))
    await page.waitForFunction(async () => { const b = await (window as unknown as { __spa: { book(): Promise<BookInput> } }).__spa.book(); return b.sheets[0]?.autoFilter?.criteria?.columns?.[0]?.filter?.type === 'top' })
    await page.locator('[data-xs-af="2"]').click()
    await page.locator('[data-xs-stored-filter]').waitFor()
    await page.locator('.xs-value-all').waitFor()
    assert.equal(await page.locator('[data-xs-apply]').isDisabled(), true)
    await page.screenshot({ path: resolve(out, `rich-filter-${size.name}-stored.png`) })
    assert.deepEqual(errors, [])
    await page.close()
    console.log(`${size.name}: run fonts, saved criteria, Apply, Clear, saved xlsx and opaque conditions passed`)
  }
} finally {
  await browser.close()
  await new Promise<void>((done, reject) => server.httpServer.close(e => e ? reject(e) : done()))
}
