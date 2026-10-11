// The app like Excel in a browser: the example's spa.html built for production and served by `vite preview`, driven by
// Playwright on a desktop (1920 × 1200, mouse and keys) and a phone (390 × 844, touch). Each creates a ru-RU book, types
// values and a formula, makes them bold and percent and draws outside borders from the toolbar, undoes and redoes, saves
// the book, opens the saved file again (dropped on the page, or from the picker) and checks values and formatting.
// On the phone every target is at least 44 px and the page never scrolls sideways. Run: npm run check:spa.
// Output: test-results/screens/spa-*.png and the saved files.
import { mkdirSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { chromium, type Browser, type Page } from 'playwright'
import { build, preview } from 'vite'
import type { BookInput, CellStyle } from 'xlcalc'

const root = resolve(import.meta.dirname, '..')
const config = resolve(root, 'vite.example.config.ts')
const out = resolve(root, 'test-results/screens')
mkdirSync(out, { recursive: true })

const THIN = { style: 'thin' }

interface Size {
  name: string
  width: number
  height: number
  phone: boolean
}

/** Wait until a cell (`row:col`) shows `text`. */
async function shows(page: Page, cell: string, text: string): Promise<void> {
  await page.waitForFunction(([c, t]) => document.querySelector(`[data-cell="${c}"]`)?.textContent?.trim() === t, [cell, text] as const, { timeout: 30_000 })
}
/** Wait for the status bar's Ready. */
async function ready(page: Page): Promise<void> {
  await page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
}
/** The book as the app's source holds it now. */
const bookOf = (page: Page) => page.evaluate(() => (window as any).__spa.book() as Promise<BookInput>)
/** A cell of the first sheet: its own style object and number format. */
function cellOf(book: BookInput, addr: string): { style: CellStyle | undefined; numberFormat: unknown; formula?: string; value?: unknown } {
  const c = (book.sheets[0]!.cells as Record<string, any>)[addr] ?? {}
  return { style: c.style === undefined ? undefined : book.styles?.[c.style], numberFormat: c.numberFormat, formula: c.formula, value: c.value }
}
/** JSON with keys in order, so that objects compare by their fields alone. */
const canon = (v: unknown): string => JSON.stringify(v, (_k, x) => (x && typeof x === 'object' && !Array.isArray(x) ? Object.fromEntries(Object.entries(x).sort(([a], [b]) => a.localeCompare(b))) : x))
function expect(what: string, got: unknown, want: unknown): void {
  if (canon(got) !== canon(want)) throw new Error(`${what}: ${canon(got)}, want ${canon(want)}`)
}
/** The outside borders of A1:A3 as the book holds them: top on A1, bottom on A3, left and right on all three. */
function bordersOf(book: BookInput): unknown[] {
  return ['A1', 'A2', 'A3'].map((a) => cellOf(book, a).style?.border ?? null)
}
const OUTSIDE = [
  { top: THIN, left: THIN, right: THIN },
  { left: THIN, right: THIN },
  { bottom: THIN, left: THIN, right: THIN },
]

/** Every button, field and list visible in the app bar and the toolbar (and its menus) at least 44 × 44 px; fields by their height. */
async function checkTargets(page: Page, where: string): Promise<void> {
  const small = await page.evaluate(() =>
    [...document.querySelectorAll<HTMLElement>('.spa-top button, .spa-top input:not([type=file]), .xs-toolbar button, .xs-toolbar select, .spa-dialog button, .spa-dialog select, .spa-choice')]
      .filter((el) => el.getClientRects().length > 0 && getComputedStyle(el).visibility !== 'hidden')
      .map((el) => {
        const r = el.getBoundingClientRect()
        return { what: `${el.tagName.toLowerCase()}[${el.dataset.tb ?? el.dataset.spa ?? el.className}]`, w: Math.round(r.width), h: Math.round(r.height), field: el.matches('input, select, .spa-choice') }
      })
      .filter((b) => b.h < 44 || (!b.field && b.w < 44)),
  )
  if (small.length) throw new Error(`${where}: touch targets under 44 px: ${JSON.stringify(small)}`)
}
async function checkNoSideScroll(page: Page, where: string): Promise<void> {
  const { scroll, width } = await page.evaluate(() => ({ scroll: Math.max(document.documentElement.scrollWidth, document.body.scrollWidth), width: window.innerWidth }))
  if (scroll > width) throw new Error(`${where}: the page scrolls sideways (${scroll} > ${width})`)
}

/** A text into a cell: typed over it on a desktop, through the formula bar on a phone. */
async function enter(page: Page, size: Size, cell: string, text: string): Promise<void> {
  if (size.phone) {
    await page.locator(`[data-cell="${cell}"]`).tap()
    await page.locator('.xs-fx').tap()
    await page.keyboard.press('ControlOrMeta+A')
  } else await page.locator(`[data-cell="${cell}"]`).click()
  await page.keyboard.type(text)
  await page.keyboard.press('Enter')
}
async function press(page: Page, size: Size, selector: string): Promise<void> {
  if (size.phone) await page.locator(selector).tap()
  else await page.locator(selector).click()
}

async function scenario(browser: Browser, base: string, size: Size): Promise<void> {
  const context = await browser.newContext({
    viewport: { width: size.width, height: size.height },
    deviceScaleFactor: size.phone ? 2 : 1,
    isMobile: size.phone,
    hasTouch: size.phone,
    acceptDownloads: true,
  })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  let step = 0
  const shot = (name: string) => page.screenshot({ path: resolve(out, `spa-${size.name}-${String(++step).padStart(2, '0')}-${name}.png`) })
  const at = (what: string) => `[${size.name}] ${what}`

  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  // The status bar's texts as they come, to see the recomputation show and go.
  await page.evaluate(() => {
    const seen = new Set<string>()
    const el = document.querySelector('[data-spa-status]')!
    new MutationObserver(() => seen.add(el.textContent!.trim())).observe(el, { childList: true, characterData: true, subtree: true })
    ;(window as any).__seen = seen
  })
  if (size.phone) await checkNoSideScroll(page, at('open'))

  // 1. A new ru-RU book.
  await press(page, size, '[data-spa="new"]')
  await page.waitForSelector('[data-spa-locale="ru-RU"]')
  if (size.phone) await checkTargets(page, at('new book dialog'))
  await press(page, size, '[data-spa-locale="ru-RU"]')
  await shot('new')
  await press(page, size, '[data-spa="create"]')
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  expect(at('a new book is ru-RU'), (await bookOf(page)).locale, 'ru-RU')

  // 2. Values and a formula in ru-RU.
  await enter(page, size, '1:1', '0,25')
  await enter(page, size, '2:1', '0,5')
  await enter(page, size, '3:1', '=СУММ(A1:A2)')
  await shows(page, '3:1', '0,75')
  await ready(page)
  await shot('values')

  // 3. A1:A3 bold: by the mouse with Shift, on a phone through the address box.
  if (size.phone) {
    await page.locator('.xs-address').tap()
    await page.locator('.xs-address').fill('A1:A3')
    await page.keyboard.press('Enter')
  } else {
    await page.locator('[data-cell="1:1"]').click()
    await page.locator('[data-cell="3:1"]').click({ modifiers: ['Shift'] })
  }
  await press(page, size, '[data-tb="bold"]')
  await page.waitForSelector('[data-tb="bold"][aria-pressed="true"]')
  const weight = await page.evaluate(() => getComputedStyle(document.querySelector('[data-cell="3:1"] .xs-in > *')!).fontWeight)
  expect(at('A3 drawn bold'), weight, '700')
  await shot('bold')

  // 4. Percent from the number format list (in the menu on a phone).
  if (size.phone) {
    await press(page, size, '[data-tb="more"]')
    await page.waitForSelector('[data-tb-menu]:not([style*="none"]) [data-tb="number"]')
    await checkTargets(page, at('toolbar menu'))
    await checkNoSideScroll(page, at('toolbar menu'))
    await shot('menu')
  }
  await page.locator('[data-tb="number"]').selectOption('percent')
  await shows(page, '1:1', '25%')
  await shows(page, '3:1', '75%')
  // The source reloads after formatting: the list shows the new format of the active cell.
  await page.waitForFunction(() => (document.querySelector('[data-tb="number"]') as HTMLSelectElement | null)?.value === 'percent', null, { timeout: 30_000 })
  await shot('percent')

  // 5. Outside borders: four edges in one step.
  await press(page, size, '[data-tb="borders"]')
  await page.waitForSelector('[data-tb="border-outside"]')
  if (size.phone) {
    await checkTargets(page, at('borders menu'))
    await checkNoSideScroll(page, at('borders menu'))
  }
  await shot('borders-menu')
  await press(page, size, '[data-tb="border-outside"]')
  await ready(page)
  await page.waitForFunction(() => getComputedStyle(document.querySelector('[data-cell="3:1"]')!).borderBottomColor === 'rgb(0, 0, 0)')
  expect(at('outside borders'), bordersOf(await bookOf(page)), OUTSIDE)
  if (size.phone) {
    await press(page, size, '[data-tb="font-palette"]')
    await checkTargets(page, at('palette'))
    await checkNoSideScroll(page, at('palette'))
    await shot('palette')
    await press(page, size, '[data-tb="font-palette"]')
    await press(page, size, '[data-tb="more"]')
  }
  await shot('outside')

  // 6. Undo takes the borders back whole, redo draws them again.
  await press(page, size, '[data-tb="undo"]')
  await ready(page)
  await page.waitForFunction(() => getComputedStyle(document.querySelector('[data-cell="3:1"]')!).borderBottomColor !== 'rgb(0, 0, 0)')
  expect(at('undone'), bordersOf(await bookOf(page)), [null, null, null])
  await shot('undo')
  await press(page, size, '[data-tb="redo"]')
  await ready(page)
  expect(at('redone'), bordersOf(await bookOf(page)), OUTSIDE)
  await shot('redo')

  // 7. Save as, under the name typed into the bar.
  await page.locator('[data-spa-name]').fill(`Проверка-${size.name}`)
  const [download] = await Promise.all([page.waitForEvent('download'), press(page, size, '[data-spa="save"]')])
  expect(at('file name'), download.suggestedFilename(), `Проверка-${size.name}.xlsx`)
  const saved = resolve(out, `spa-${size.name}-saved.xlsx`)
  await download.saveAs(saved)
  await ready(page)
  await shot('saved')

  // 8. Open the saved file: dropped on the page on a desktop, from the picker on a phone.
  // A cell typed after saving (Latin: Playwright types Cyrillic without keydown, which the grid needs to open its field).
  await enter(page, size, '1:2', 'draft')
  await shows(page, '1:2', 'draft')
  const file = { name: download.suggestedFilename(), mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await readFile(saved) }
  if (size.phone) await page.locator('[data-spa-file]').setInputFiles(file)
  else {
    const drop = await page.evaluateHandle(
      ([data, name, type]) => {
        const dt = new DataTransfer()
        dt.items.add(new File([new Uint8Array(data)], name, { type }))
        return dt
      },
      [[...file.buffer], file.name, file.mimeType] as const,
    )
    await page.dispatchEvent('.spa', 'dragover', { dataTransfer: drop })
    await page.waitForSelector('.spa-drop')
    await shot('drop')
    await page.dispatchEvent('.spa', 'drop', { dataTransfer: drop })
  }
  await shows(page, '1:1', '25%')
  await ready(page)

  // 9. Values and formatting as saved; the cell typed after saving is not there.
  for (const [cell, text] of [['1:1', '25%'], ['2:1', '50%'], ['3:1', '75%'], ['1:2', '']] as const) await shows(page, cell, text)
  expect(at('name of the opened book'), await page.locator('[data-spa-name]').inputValue(), `Проверка-${size.name}`)
  const book = await bookOf(page)
  expect(at('reopened locale'), book.locale, 'ru-RU')
  expect(at('reopened A3'), { formula: cellOf(book, 'A3').formula, bold: cellOf(book, 'A3').style?.font?.bold, format: cellOf(book, 'A3').numberFormat }, { formula: 'SUM(A1:A2)', bold: true, format: 9 })
  expect(at('reopened borders'), bordersOf(book), OUTSIDE)
  await page.locator('[data-cell="2:1"]').click()
  await page.waitForSelector('[data-tb="bold"][aria-pressed="true"]')
  expect(at('the list shows the opened percent'), await page.locator('[data-tb="number"]').inputValue(), 'percent')
  const seen = await page.evaluate(() => [...((window as any).__seen as Set<string>)])
  if (!seen.includes('Пересчёт…')) throw new Error(`${at('status')}: never showed the recomputation: ${JSON.stringify(seen)}`)
  if (size.phone) {
    await checkTargets(page, at('reopened'))
    await checkNoSideScroll(page, at('reopened'))
  }
  await shot('reopened')
  const t = await page.evaluate(() => (window as any).__spa.timings() as { open: number; change: number })
  console.log(`${at('ok')}: new ru-RU book, values, bold, percent, outside borders, undo, redo, saved as ${download.suggestedFilename()}, opened again in ${Math.round(t.open)} ms`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

await build({ configFile: config, logLevel: 'error' })
const server = await preview({ configFile: config, preview: { port: 0 }, logLevel: 'error' })
const base = server.resolvedUrls!.local[0]!
const browser = await chromium.launch()
try {
  await scenario(browser, base, { name: '1920', width: 1920, height: 1200, phone: false })
  await scenario(browser, base, { name: '390', width: 390, height: 844, phone: true })
} finally {
  await browser.close()
  await new Promise<void>((done) => server.httpServer.close(() => done()))
}
console.log(`screens in ${out}`)
