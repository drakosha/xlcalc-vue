// The app like Excel in a browser: the example's spa.html built for production and served by `vite preview`, driven by
// Playwright on a desktop (1920 × 1200, mouse and keys) and a phone (390 × 844, touch). Each creates a ru-RU book, types
// values and a formula, makes them bold and percent and draws outside borders from the toolbar, undoes and redoes, saves
// the book, opens the saved file again (dropped on the page, or from the picker) and checks values and formatting; on
// the desktop a column in percent shows its format in the list on its empty cell.
// Then the structure menu on a new book: on the desktop a row inserted by a right click (the formula under it moves),
// a column hidden and shown, a sheet added by "+", renamed and deleted, a column fitted by a double click on its edge;
// on the phone a row inserted from the menu a long press opens.
// Then Find and Replace and Format Cells on a new book: a search with Match case and in values, Replace All and its
// Control+Z, the dialog's font, outline and percent with OK and their undo.
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

/** Every button, field and list visible in the app bar, the toolbar (and its menus), the find bar and the dialogs at least 44 × 44 px; fields by their height. */
async function checkTargets(page: Page, where: string): Promise<void> {
  const small = await page.evaluate(() =>
    [
      ...document.querySelectorAll<HTMLElement>(
        '.spa-top button, .spa-top input:not([type=file]), .xs-toolbar button, .xs-toolbar select, .spa-dialog button, .spa-dialog select, .spa-choice, .xs-menu button, .xs-dialog button, .xs-dialog input, .xs-tabs button, ' +
          '.xs-find button, .xs-find-input, .xs-find-select, .xs-find-opt, .xs-fmt button, .xs-fmt-input, .xs-fmt-select, .xs-fmt-choice, .xs-fmt-check',
      ),
    ]
      .filter((el) => el.getClientRects().length > 0 && getComputedStyle(el).visibility !== 'hidden')
      .map((el) => {
        const r = el.getBoundingClientRect()
        return {
          what: `${el.tagName.toLowerCase()}[${el.dataset.tb ?? el.dataset.spa ?? el.dataset.xsMenu ?? el.className}]`,
          w: Math.round(r.width),
          h: Math.round(r.height),
          field: el.matches('input, select, .spa-choice, .xs-tab, .xs-find-opt, .xs-fmt-choice, .xs-fmt-check'),
        }
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
  if (!size.phone) {
    // 10. Column C in percent: its empty cell shows the column's format in the list, as Excel's does.
    await page.locator('.xs-head[data-col="3"]').click()
    await page.locator('[data-tb="number"]').selectOption('percent')
    await ready(page)
    const formatted = await bookOf(page)
    expect(at('column C in percent'), (formatted.sheets[0]!.columns as Record<string, any>).C.numberFormat, 9)
    expect(at('C7 has no format of its own'), cellOf(formatted, 'C7').numberFormat, undefined)
    await page.locator('[data-cell="7:3"]').click()
    await page.waitForFunction(() => (document.querySelector('[data-tb="number"]') as HTMLSelectElement | null)?.value === 'percent', null, { timeout: 30_000 })
    await page.locator('[data-cell="7:4"]').click()
    await page.waitForFunction(() => (document.querySelector('[data-tb="number"]') as HTMLSelectElement | null)?.value === 'general', null, { timeout: 30_000 })
    await page.locator('[data-cell="7:3"]').click()
    await shot('column-format')
  }
  const t = await page.evaluate(() => (window as any).__spa.timings() as { open: number; change: number })
  console.log(`${at('ok')}: new ru-RU book, values, bold, percent, outside borders, undo, redo, saved as ${download.suggestedFilename()}, opened again in ${Math.round(t.open)} ms${size.phone ? '' : ', a column\'s format on its empty cell'}`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

/** A finger held still on an element for 0.7 s, as Chromium's touch input gets it. */
async function longPress(page: Page, selector: string): Promise<void> {
  const box = (await page.locator(selector).boundingBox())!
  const point = { x: box.x + box.width / 2, y: box.y + box.height / 2 }
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [point] })
  await page.waitForTimeout(700)
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] })
  await cdp.detach()
}
/** The letters of the column headers the grid draws, in order. */
const letters = (page: Page) => page.evaluate(() => [...document.querySelectorAll('.xs-head .xs-head-label')].map((h) => h.textContent!.trim()))
const tabNames = (page: Page) => page.evaluate(() => [...document.querySelectorAll('.xs-tab')].map((t) => t.textContent!.trim()))
const activeTab = (page: Page) => page.evaluate(() => document.querySelector('.xs-tab-on')?.textContent?.trim())

/** The structure menu on a new ru-RU book: rows, columns, sheets and the fit by a double click (on a phone a long press). */
async function structureScenario(browser: Browser, base: string, size: Size): Promise<void> {
  const context = await browser.newContext({ viewport: { width: size.width, height: size.height }, deviceScaleFactor: size.phone ? 2 : 1, isMobile: size.phone, hasTouch: size.phone })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  let step = 0
  const shot = (name: string) => page.screenshot({ path: resolve(out, `spa-${size.name}-s${String(++step).padStart(2, '0')}-${name}.png`) })
  const at = (what: string) => `[${size.name} structure] ${what}`

  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  await press(page, size, '[data-spa="new"]')
  await press(page, size, '[data-spa-locale="ru-RU"]')
  await press(page, size, '[data-spa="create"]')
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  await enter(page, size, '1:1', '1')
  await enter(page, size, '2:1', '2')
  await enter(page, size, '3:1', '=СУММ(A1:A2)')
  await shows(page, '3:1', '3')
  await ready(page)
  // One sheet has its tab and "+", as in Excel.
  expect(at('one sheet has its tab'), await tabNames(page), ['Лист1'])

  // 1. A row inserted above row 2 from the menu: the sum moves to A4 and reads A1:A3.
  if (size.phone) await longPress(page, '.xs-num[data-row="2"]')
  else await page.locator('.xs-num[data-row="2"]').click({ button: 'right' })
  await page.waitForSelector('[data-xs-menu="insert"]')
  if (size.phone) {
    await checkTargets(page, at('menu'))
    await checkNoSideScroll(page, at('menu'))
  }
  await shot('row-menu')
  await press(page, size, '[data-xs-menu="insert"]')
  await shows(page, '4:1', '3')
  await ready(page)
  expect(at('the sum moved'), cellOf(await bookOf(page), 'A4').formula, 'SUM(A1:A3)')
  await shot('row-inserted')
  if (size.phone) {
    // The width box of a letter's menu and a tab's menu, by finger too.
    await longPress(page, '.xs-head[data-col="2"]')
    await press(page, size, '[data-xs-menu="size"]')
    await page.waitForSelector('[data-xs-dialog-input]')
    await checkTargets(page, at('width box'))
    await checkNoSideScroll(page, at('width box'))
    await shot('width-box')
    await press(page, size, '[data-xs-dialog-cancel]')
    await longPress(page, '.xs-tab')
    await page.waitForSelector('[data-xs-menu="rename"]')
    await checkTargets(page, at('tab menu'))
    await checkNoSideScroll(page, at('tab menu'))
    await shot('tab-menu')
    await page.locator('[data-cell="6:3"]').tap()
    await page.waitForFunction(() => !document.querySelector('[data-xs-menu]'))
    console.log(`${at('ok')}: a row inserted from the menu a long press opened; the width box and a tab's menu by finger`)
    if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
    await context.close()
    return
  }

  // 2. Column B hidden from its letter's menu, shown again over A:C.
  await page.locator('.xs-head[data-col="2"]').click({ button: 'right' })
  await page.locator('[data-xs-menu="hide"]').click()
  await page.waitForFunction(() => !document.querySelector('.xs-head[data-col="2"]'))
  await ready(page)
  expect(at('B hidden'), (await letters(page)).slice(0, 3), ['A', 'C', 'D'])
  await shot('column-hidden')
  await page.locator('.xs-head[data-col="1"]').click()
  await page.locator('.xs-head[data-col="3"]').click({ modifiers: ['Shift'] })
  await page.locator('.xs-head[data-col="1"]').click({ button: 'right' })
  await page.locator('[data-xs-menu="unhide"]').click()
  await page.waitForSelector('.xs-head[data-col="2"]')
  await ready(page)
  expect(at('B shown'), (await letters(page)).slice(0, 3), ['A', 'B', 'C'])

  // 3. A sheet added by "+", renamed on its tab, deleted: the sheet before it shows again.
  await page.locator('[data-xs-add-sheet]').click()
  await page.waitForFunction(() => document.querySelector('.xs-tab-on')?.textContent?.trim() === 'Лист2')
  await ready(page)
  await page.locator('.xs-tab-on').click({ button: 'right' })
  await page.locator('[data-xs-menu="rename"]').click()
  await page.locator('.xs-tab-input').fill('Totals')
  await shot('rename')
  await page.keyboard.press('Enter')
  await page.waitForFunction(() => document.querySelector('.xs-tab-on')?.textContent?.trim() === 'Totals')
  await ready(page)
  expect(at('tabs after rename'), await tabNames(page), ['Лист1', 'Totals'])
  await page.locator('.xs-tab-on').click({ button: 'right' })
  await shot('sheet-menu')
  await page.locator('[data-xs-menu="deleteSheet"]').click()
  await page.waitForFunction(() => document.querySelectorAll('.xs-tab').length === 1)
  await ready(page)
  expect(at('after delete'), [await tabNames(page), await activeTab(page)], [['Лист1'], 'Лист1'])

  // 4. A long text in B1, then a double click on B's edge fits B to it.
  await enter(page, size, '1:2', 'a much longer text than the column')
  await shows(page, '1:2', 'a much longer text than the column')
  await ready(page)
  const widthOf = () => page.evaluate(() => document.querySelector<HTMLElement>('.xs-head[data-col="2"]')!.getBoundingClientRect().width)
  const before = await widthOf()
  await page.locator('.xs-head[data-col="2"] .xs-grip').dblclick()
  await page.waitForFunction((w) => document.querySelector<HTMLElement>('.xs-head[data-col="2"]')!.getBoundingClientRect().width > w + 40, before)
  await ready(page)
  const fitted = (await bookOf(page)).sheets[0]!.columns?.B?.width
  if (!(typeof fitted === 'number' && fitted > 20)) throw new Error(`${at('fit')}: B is ${fitted} wide in the book`)
  await shot('fitted')
  console.log(`${at('ok')}: row inserted (A4 is SUM(A1:A3)), B hidden and shown, sheet added, renamed and deleted, B fitted from ${Math.round(before)} px to ${Math.round(await widthOf())} px (${fitted})`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

/** The find bar's status. */
const findStatus = (page: Page) => page.evaluate(() => document.querySelector('.xs-find-status')?.textContent?.trim() ?? '')
async function statusIs(page: Page, text: string, where: string): Promise<void> {
  await page.waitForFunction((t) => document.querySelector('.xs-find-status')?.textContent?.trim() === t, text, { timeout: 30_000 }).catch(async () => {
    throw new Error(`${where}: the find bar says ${JSON.stringify(await findStatus(page))}, want ${JSON.stringify(text)}`)
  })
}

/**
 * Find and Replace and Format Cells on a new ru-RU book: a search with Match case, then in values, Replace All and its
 * Control+Z, and the dialog's font, outline and percent over A1:B3 with OK, one step of undo.
 */
async function findFormatScenario(browser: Browser, base: string, size: Size): Promise<void> {
  const context = await browser.newContext({ viewport: { width: size.width, height: size.height }, deviceScaleFactor: size.phone ? 2 : 1, isMobile: size.phone, hasTouch: size.phone })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  let step = 0
  const shot = (name: string) => page.screenshot({ path: resolve(out, `spa-${size.name}-f${String(++step).padStart(2, '0')}-${name}.png`) })
  const at = (what: string) => `[${size.name} find and format] ${what}`

  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  await press(page, size, '[data-spa="new"]')
  await press(page, size, '[data-spa-locale="ru-RU"]')
  await press(page, size, '[data-spa="create"]')
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  for (const [cell, text] of [['1:1', 'tea'], ['2:1', 'coffee'], ['3:1', 'Tea time'], ['4:1', '=A1&" cup"'], ['1:2', '0,25'], ['2:2', '0,5']] as const) await enter(page, size, cell, text)
  await shows(page, '4:1', 'tea cup')
  await ready(page)

  // 1. Find: "tea" in any case, then with Match case, then in values, where A4 shows it by its formula.
  if (size.phone) await press(page, size, '[data-xs-find]')
  else {
    await page.locator('[data-cell="5:1"]').click()
    await page.keyboard.press('Control+f')
  }
  await page.locator('.xs-find-input').fill('tea')
  await statusIs(page, '1 из 2', at('tea'))
  await press(page, size, '[data-xs-find-options]')
  await press(page, size, '[data-xs-find-case]')
  await statusIs(page, '1 из 1', at('tea, match case'))
  if (size.phone) {
    await checkTargets(page, at('find options'))
    await checkNoSideScroll(page, at('find options'))
  }
  await shot('find-case')
  await page.locator('[data-xs-find-look]').selectOption('values')
  await statusIs(page, '2 из 2', at('tea, match case, in values'))
  expect(at('the hit in values'), await page.evaluate(() => document.querySelector('.xs-cell.xs-active')?.getAttribute('data-cell')), '4:1')
  await shot('find-values')

  // 2. Replace All in formulas: tea in any case to milk, A4 computes from the new A1; Control+Z takes it back whole.
  await press(page, size, '[data-xs-find-case]')
  await press(page, size, '[data-xs-find-replace]')
  await page.locator('[data-xs-replace-input]').fill('milk')
  await statusIs(page, '1 из 2', at('tea, before Replace All'))
  if (size.phone) {
    await checkTargets(page, at('replace row'))
    await checkNoSideScroll(page, at('replace row'))
  }
  await shot('replace-row')
  await press(page, size, '[data-xs-replace-all]')
  await statusIs(page, 'Сделано замен: 2', at('Replace All'))
  for (const [cell, text] of [['1:1', 'milk'], ['3:1', 'milk time'], ['4:1', 'milk cup']] as const) await shows(page, cell, text)
  await ready(page)
  expect(at('A3 after Replace All'), cellOf(await bookOf(page), 'A3').value, 'milk time')
  await shot('replaced')
  await page.locator('[data-cell="6:3"]').click()
  await page.keyboard.press('Control+z')
  for (const [cell, text] of [['1:1', 'tea'], ['3:1', 'Tea time'], ['4:1', 'tea cup']] as const) await shows(page, cell, text)
  await ready(page)
  expect(at('A1 after Control+Z'), cellOf(await bookOf(page), 'A1').value, 'tea')
  await shot('replace-undone')
  await page.locator('.xs-find [aria-label="Закрыть"]').last().click()

  // 3. Format Cells over A1:B3 from the toolbar: bold 14 pt, the outline, percent; OK formats it in one step of undo.
  if (size.phone) {
    await page.locator('.xs-address').tap()
    await page.locator('.xs-address').fill('A1:B3')
    await page.keyboard.press('Enter')
    await press(page, size, '[data-tb="more"]')
  } else {
    await page.locator('[data-cell="1:1"]').click()
    await page.locator('[data-cell="3:2"]').click({ modifiers: ['Shift'] })
  }
  await press(page, size, '[data-tb="format-cells"]')
  await page.waitForSelector('[data-fmt]')
  await press(page, size, '[data-fmt-tab="font"]')
  await page.locator('[data-fmt-font-style]').selectOption('bold')
  await page.locator('[data-fmt-font-size]').fill('14')
  if (size.phone) {
    await checkTargets(page, at('dialog, font'))
    await checkNoSideScroll(page, at('dialog, font'))
  }
  await shot('dialog-font')
  await press(page, size, '[data-fmt-tab="border"]')
  await press(page, size, '[data-fmt-line="medium"]')
  await press(page, size, '[data-fmt-preset="outline"]')
  if (size.phone) {
    await checkTargets(page, at('dialog, border'))
    await checkNoSideScroll(page, at('dialog, border'))
  }
  await shot('dialog-border')
  await press(page, size, '[data-fmt-tab="number"]')
  await press(page, size, '[data-fmt-number="percent"]')
  await shot('dialog-number')
  await press(page, size, '[data-fmt-ok]')
  await page.waitForSelector('[data-fmt]', { state: 'detached' })
  await shows(page, '1:2', '25%')
  await ready(page)
  const MEDIUM = { style: 'medium' }
  const formatted = await bookOf(page)
  expect(
    at('A1 and B3 formatted'),
    [cellOf(formatted, 'A1').style, cellOf(formatted, 'B3').style?.border, cellOf(formatted, 'B1').numberFormat],
    [{ font: { bold: true, size: 14 }, border: { top: MEDIUM, left: MEDIUM } }, { bottom: MEDIUM, right: MEDIUM }, 9],
  )
  await shot('formatted')
  if (size.phone) await press(page, size, '[data-tb="undo"]')
  else await page.keyboard.press('Control+z')
  await shows(page, '1:2', '0,25')
  await ready(page)
  const undone = await bookOf(page)
  expect(at('Format Cells undone in one step'), [cellOf(undone, 'A1').style, cellOf(undone, 'B1').numberFormat], [undefined, undefined])
  await shot('format-undone')
  console.log(`${at('ok')}: found with Match case and in values, Replace All of 2 cells and its Control+Z, Format Cells (bold 14, outline, percent) and its undo`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

await build({ configFile: config, logLevel: 'error' })
const server = await preview({ configFile: config, preview: { port: 0 }, logLevel: 'error' })
const base = server.resolvedUrls!.local[0]!
const browser = await chromium.launch()
try {
  const desktop = { name: '1920', width: 1920, height: 1200, phone: false }
  const phone = { name: '390', width: 390, height: 844, phone: true }
  await scenario(browser, base, desktop)
  await scenario(browser, base, phone)
  await structureScenario(browser, base, desktop)
  await structureScenario(browser, base, phone)
  await findFormatScenario(browser, base, desktop)
  await findFormatScenario(browser, base, phone)
} finally {
  await browser.close()
  await new Promise<void>((done) => server.httpServer.close(() => done()))
}
console.log(`screens in ${out}`)
