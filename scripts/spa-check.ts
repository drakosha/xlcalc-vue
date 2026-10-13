// The app like Excel in a browser: the example's spa.html built for production and served by `vite preview`, driven by
// Playwright on a desktop (1920 × 1200, mouse and keys) and a phone (390 × 844, touch). Each creates a ru-RU book, types
// values and a formula, makes them bold and Percentage (Excel's 0.00%) and draws outside borders from the toolbar, undoes and redoes, saves
// the book, opens the saved file again (dropped on the page, or from the picker) and checks values and formatting; on
// the desktop a column in percent shows its format's category in the box on its empty cell.
// Then the structure menu on a new book: on the desktop a row inserted by a right click (the formula under it moves),
// a column hidden and shown, a sheet added by "+", renamed and deleted, a column fitted by a double click on its edge;
// on the phone a row inserted from the menu a long press opens.
// Then Find and Replace and Format Cells on a new book: a search with Match case and in values, Replace All and its
// Control+Z, the dialog's font, outline and percent with OK and their undo.
// Then the fill handle and cut, copy and paste inside the book: a formula dragged three rows down by its handle, a range
// with a formula copied and pasted, a formula cut and pasted, each undone (Control+C, X, V on the desktop, the cells' menu
// on the phone).
// Then Formula AutoComplete: `=СУ` typed into a cell (the formula bar on the phone) lists СУММ first, Tab (a tap on the
// phone) enters `СУММ(`, and `A1:A3)` with Enter sums the column.
// Then, on the desktop, a file with a circular reference opened from the picker holds the values it stores, as Excel
// opens it, through an edit elsewhere; Control+Down goes to the end of a block (Range.End); and a cell copied over a
// whole column is refused at once by the app's limit of cells one change writes, the book as it was.
// On the phone every target is at least 44 px and the page never scrolls sideways. Run: npm run check:spa.
// DEMO=1 npm run check:spa runs it on the demo's build (`--mode demo`, served under /demo/), which /demo/ opens and
// which holds no source map.
// Output: test-results/screens/spa-*.png and the saved files.
import { mkdirSync, readdirSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { chromium, type Browser, type Page } from 'playwright'
import { build, preview } from 'vite'
import type { BookInput, CellStyle } from 'xlcalc'
import { parseCell, parseRange } from 'xlcalc/a1'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { numberFormatCategory } from 'xlcalc/text'
import { writeXlsx } from 'xlcalc/xlsx'

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
/** The name ru-RU's box gives Percentage (`0.00%`), by xlcalc's rule, whichever Excel's the library follows. */
const PERCENT = numberFormatCategory(10, ruRU).name
/** Wait until the toolbar's Number Format box names `name`. */
async function boxNames(page: Page, name: string, where: string): Promise<void> {
  const shown = () => page.evaluate(() => (document.querySelector('[data-tb="number"]') as HTMLSelectElement | null)?.selectedOptions[0]?.textContent?.trim())
  await page.waitForFunction((n) => (document.querySelector('[data-tb="number"]') as HTMLSelectElement | null)?.selectedOptions[0]?.textContent?.trim() === n, name, { timeout: 30_000 }).catch(async () => {
    throw new Error(`${where}: the box names ${JSON.stringify(await shown())}, want ${JSON.stringify(name)}`)
  })
}
/** Wait for the status bar's Ready. */
async function ready(page: Page): Promise<void> {
  await page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
}
/** The book as the app's source holds it now. */
const bookOf = (page: Page) => page.evaluate(() => (window as any).__spa.book() as Promise<BookInput>)
/** The last look of a sheet's look layer over a cell: what an empty cell with only a look of its own shows (xlcalc 0.28). */
function lookOf(sheet: BookInput['sheets'][number], addr: string): { style?: number; numberFormat?: unknown } | undefined {
  const at = parseCell(addr)!
  const looks = sheet.looks ?? []
  for (let i = looks.length - 1; i >= 0; i--) {
    const r = parseRange(looks[i]!.range)
    if (r && at.row >= r.top && at.row <= r.bottom && at.col >= r.left && at.col <= r.right) return looks[i]
  }
  return undefined
}
/** A cell of the first sheet: its own style object and number format, from its entry or else the sheet's look layer. */
function cellOf(book: BookInput, addr: string): { style: CellStyle | undefined; numberFormat: unknown; formula?: string; value?: unknown } {
  const c = (book.sheets[0]!.cells as Record<string, any>)[addr] ?? lookOf(book.sheets[0]!, addr) ?? {}
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
        '.spa-top button, .spa-top input:not([type=file]), .xs-toolbar button, .xs-toolbar select, .spa-dialog button, .spa-dialog select, .spa-choice, .xs-menu button, .xs-dialog button, .xs-dialog input, .xs-tabs button, .xs-bar button, ' +
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
  await page.locator('[data-tb="number"]').selectOption('percentage')
  await shows(page, '1:1', '25,00%')
  await shows(page, '3:1', '75,00%')
  // The source reloads after formatting: the list shows the new format of the active cell.
  await boxNames(page, PERCENT, at('the box after Percentage'))
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
  if (size.phone) {
    const [picker] = await Promise.all([page.waitForEvent('filechooser'), press(page, size, '[data-spa="open"]')])
    await picker.setFiles(file)
  } else {
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
  await shows(page, '1:1', '25,00%')
  await ready(page)

  // 9. Values and formatting as saved; the cell typed after saving is not there.
  for (const [cell, text] of [['1:1', '25,00%'], ['2:1', '50,00%'], ['3:1', '75,00%'], ['1:2', '']] as const) await shows(page, cell, text)
  expect(at('name of the opened book'), await page.locator('[data-spa-name]').inputValue(), `Проверка-${size.name}`)
  const book = await bookOf(page)
  expect(at('reopened locale'), book.locale, 'ru-RU')
  expect(at('reopened A3'), { formula: cellOf(book, 'A3').formula, bold: cellOf(book, 'A3').style?.font?.bold, format: cellOf(book, 'A3').numberFormat }, { formula: 'SUM(A1:A2)', bold: true, format: 10 })
  expect(at('reopened borders'), bordersOf(book), OUTSIDE)
  await page.locator('[data-cell="2:1"]').click()
  await page.waitForSelector('[data-tb="bold"][aria-pressed="true"]')
  await boxNames(page, PERCENT, at('the box on the opened percent'))
  const seen = await page.evaluate(() => [...((window as any).__seen as Set<string>)])
  if (!seen.includes('Пересчёт…')) throw new Error(`${at('status')}: never showed the recomputation: ${JSON.stringify(seen)}`)
  if (size.phone) {
    await checkTargets(page, at('reopened'))
    await checkNoSideScroll(page, at('reopened'))
  }
  await shot('reopened')
  if (!size.phone) {
    // 10. Column C in percent: its empty cell shows the column's format's category in the box, as Excel's does.
    await page.locator('.xs-head[data-col="3"]').click()
    await page.locator('[data-tb="number"]').selectOption('percentage')
    await ready(page)
    const formatted = await bookOf(page)
    expect(at('column C in percent'), (formatted.sheets[0]!.columns as Record<string, any>).C.numberFormat, 10)
    expect(at('C7 has no format of its own'), cellOf(formatted, 'C7').numberFormat, undefined)
    await page.locator('[data-cell="7:3"]').click()
    await boxNames(page, PERCENT, at('the box on C7'))
    await page.locator('[data-cell="7:4"]').click()
    await boxNames(page, numberFormatCategory(0, ruRU).name, at('the box on D7'))
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
  await press(page, size, '[data-fmt-number="percentage"]')
  await shot('dialog-number')
  await press(page, size, '[data-fmt-ok]')
  await page.waitForSelector('[data-fmt]', { state: 'detached' })
  await shows(page, '1:2', '25,00%')
  await ready(page)
  const MEDIUM = { style: 'medium' }
  const formatted = await bookOf(page)
  expect(
    at('A1 and B3 formatted'),
    [cellOf(formatted, 'A1').style, cellOf(formatted, 'B3').style?.border, cellOf(formatted, 'B1').numberFormat],
    [{ font: { bold: true, size: 14 }, border: { top: MEDIUM, left: MEDIUM } }, { bottom: MEDIUM, right: MEDIUM }, 10],
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

/** The middle of an element on the page. */
async function middle(page: Page, selector: string): Promise<{ x: number; y: number }> {
  const box = (await page.locator(selector).boundingBox())!
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 }
}
/** A finger put down at `from`, moved to `to` in steps and lifted, as Chromium's touch input gets it; `held` runs before the lift. */
async function touchDrag(page: Page, from: { x: number; y: number }, to: { x: number; y: number }, held: () => Promise<unknown>): Promise<void> {
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [from] })
  for (let i = 1; i <= 8; i++) await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: from.x + ((to.x - from.x) * i) / 8, y: from.y + ((to.y - from.y) * i) / 8 }] })
  await held()
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] })
  await cdp.detach()
}
const addressOf = (page: Page) => page.locator('.xs-address').inputValue()
const antsShown = async (page: Page) => (await page.locator('[data-xs-ants]').count()) > 0
/** A cell's formula and value in the book as the app's source holds it. */
const formulas = async (page: Page, cells: string[]) => {
  const book = await bookOf(page)
  return cells.map((a) => cellOf(book, a).formula ?? cellOf(book, a).value ?? null)
}

/**
 * The fill handle and cut, copy and paste inside the book on a new ru-RU book: B1's formula dragged three rows down by
 * its handle (the mouse, or a finger on the phone), A1:B1 copied and pasted at D6, B1 cut and pasted at D2 (Control+C, X
 * and V on the desktop, the cells' menu a long press opens on the phone), each taken back by Control+Z (the toolbar's
 * undo on the phone). The book is checked after each step, the grid too.
 */
async function fillCopyScenario(browser: Browser, base: string, size: Size): Promise<void> {
  const context = await browser.newContext({
    viewport: { width: size.width, height: size.height },
    deviceScaleFactor: size.phone ? 2 : 1,
    isMobile: size.phone,
    hasTouch: size.phone,
    permissions: ['clipboard-read', 'clipboard-write'],
  })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  let step = 0
  const shot = (name: string) => page.screenshot({ path: resolve(out, `spa-${size.name}-c${String(++step).padStart(2, '0')}-${name}.png`) })
  const at = (what: string) => `[${size.name} fill and copy] ${what}`
  /** Control+Z, or the toolbar's undo on the phone. */
  const undo = async () => {
    if (size.phone) await press(page, size, '[data-tb="undo"]')
    else await page.keyboard.press('ControlOrMeta+z')
    await ready(page)
  }
  /** Open the cells' menu on a cell (a right click, a long press) and pick an item. */
  const menuItem = async (cell: string, item: string, check = false) => {
    if (size.phone) await longPress(page, `[data-cell="${cell}"]`)
    else await page.locator(`[data-cell="${cell}"]`).click({ button: 'right' })
    await page.waitForSelector(`[data-xs-menu="${item}"]`)
    if (check) {
      await checkTargets(page, at('cells menu'))
      await checkNoSideScroll(page, at('cells menu'))
      await shot('cells-menu')
    }
    await press(page, size, `[data-xs-menu="${item}"]`)
    await page.waitForFunction(() => !document.querySelector('[data-xs-menu]'))
  }
  /** Select a cell or a range: by the mouse, or through the address box on the phone (a second tap on a cell opens its field). */
  /** After the undo of a copy onto empty cells the grid has them empty: `./local` reads the sheet again after a structure step. */
  const stale = async (cells: string[], what: string) => {
    const shown = await page.evaluate((list) => list.map((c) => document.querySelector(`[data-cell="${c}"]`)?.textContent?.trim() ?? ''), cells)
    expect(`${at('grid')} after ${what}`, shown, cells.map(() => ''))
  }
  const select = async (from: string, to: string | null = null) => {
    if (size.phone) {
      const a1 = (cell: string) => {
        const [r, c] = cell.split(':').map(Number)
        return `${String.fromCharCode(64 + c!)}${r}`
      }
      await page.locator('.xs-address').tap()
      await page.locator('.xs-address').fill(to ? `${a1(from)}:${a1(to)}` : a1(from))
      await page.keyboard.press('Enter')
    } else {
      await page.locator(`[data-cell="${from}"]`).click()
      if (to) await page.locator(`[data-cell="${to}"]`).click({ modifiers: ['Shift'] })
    }
  }

  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  await press(page, size, '[data-spa="new"]')
  await press(page, size, '[data-spa-locale="ru-RU"]')
  await press(page, size, '[data-spa="create"]')
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  for (const [cell, text] of [['1:1', '1'], ['2:1', '2'], ['3:1', '3'], ['4:1', '4'], ['1:2', '=A1*10']] as const) await enter(page, size, cell, text)
  await shows(page, '1:2', '10')
  await ready(page)

  // 1. B1's fill handle dragged to B4: the formula is copied down, each copy reading its own row; B1:B4 is selected.
  await select('1:2')
  await page.waitForSelector('[data-xs-fill-handle]')
  const handle = await middle(page, '[data-xs-fill-handle]')
  if (size.phone) {
    const box = (await page.locator('[data-xs-fill-handle]').boundingBox())!
    if (box.width < 44 || box.height < 44) throw new Error(`${at('fill handle')}: its target is ${box.width} × ${box.height} px`)
  }
  const target = await middle(page, '[data-cell="4:2"]')
  const dragged = async () => {
    await page.waitForSelector('[data-xs-fill-frame]')
    await shot('fill-drag')
  }
  if (size.phone) await touchDrag(page, handle, target, dragged)
  else {
    await page.mouse.move(handle.x, handle.y)
    await page.mouse.down()
    await page.mouse.move(target.x, target.y, { steps: 8 })
    await dragged()
    await page.mouse.up()
  }
  for (const [cell, text] of [['2:2', '20'], ['3:2', '30'], ['4:2', '40']] as const) await shows(page, cell, text)
  await ready(page)
  expect(at('the filled formulas'), await formulas(page, ['B2', 'B3', 'B4']), ['A2*10', 'A3*10', 'A4*10'])
  expect(at('the selection after the fill'), await addressOf(page), 'B1:B4')
  await shot('filled')
  await undo()
  expect(at('the fill undone'), await formulas(page, ['B2', 'B3', 'B4']), [null, null, null])
  await stale(['2:2', '3:2', '4:2'], "the fill's undo")
  await shot('fill-undone')

  // 2. A1:B1 copied: its moving border; pasted at D6, where E6 reads D6; the border stays until Esc.
  await select('1:1', '1:2')
  if (size.phone) await menuItem('1:1', 'copy', true)
  else await page.keyboard.press('ControlOrMeta+c')
  await page.waitForSelector('[data-xs-ants]')
  await shot('copied')
  await select('6:4')
  if (size.phone) await menuItem('6:4', 'paste')
  else await page.keyboard.press('ControlOrMeta+v')
  await shows(page, '6:4', '1')
  await shows(page, '6:5', '10')
  await ready(page)
  expect(at('the pasted copy'), await formulas(page, ['D6', 'E6']), [1, 'D6*10'])
  expect(at('the selection after the paste'), await addressOf(page), 'D6:E6')
  expect(at('the border stays after a paste of a copy'), await antsShown(page), true)
  await shot('pasted')
  await undo()
  expect(at('the paste undone'), await formulas(page, ['D6', 'E6']), [null, null])
  expect(at('an undo ends the copy mode'), await antsShown(page), false)
  await stale(['6:4', '6:5'], "the paste's undo")
  await shot('paste-undone')

  // 3. B1 cut and pasted at D2: the formula moves and still reads A1, B1 is empty, the border goes; the undo puts it back.
  await select('1:2')
  if (size.phone) await menuItem('1:2', 'cut')
  else await page.keyboard.press('ControlOrMeta+x')
  await page.waitForSelector('[data-xs-ants]')
  await select('2:4')
  if (size.phone) await menuItem('2:4', 'paste')
  else await page.keyboard.press('ControlOrMeta+v')
  await shows(page, '2:4', '10')
  await shows(page, '1:2', '')
  await ready(page)
  expect(at('the moved formula'), await formulas(page, ['B1', 'D2']), [null, 'A1*10'])
  expect(at('the border after a paste of a cut'), await antsShown(page), false)
  await shot('moved')
  await undo()
  await shows(page, '1:2', '10')
  await shows(page, '2:4', '')
  expect(at('the move undone'), await formulas(page, ['B1', 'D2']), ['A1*10', null])
  await shot('move-undone')
  if (size.phone) await checkNoSideScroll(page, at('after the moves'))
  console.log(`${at('ok')}: B1 filled to B4 by ${size.phone ? 'a finger' : 'the mouse'} (B4 is A4*10), A1:B1 pasted at D6 (E6 is D6*10), B1 moved to D2 (A1*10), each undone`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

/**
 * Formula AutoComplete on a new ru-RU book: `=СУ` typed into A4 (in the cell on the desktop, in the formula bar on the
 * phone) lists the functions it begins with СУММ highlighted first; Tab, or a tap on the phone, enters `СУММ(`, and
 * `A1:A3)` with Enter gives the sum of A1:A3.
 */
async function autocompleteScenario(browser: Browser, base: string, size: Size): Promise<void> {
  const context = await browser.newContext({ viewport: { width: size.width, height: size.height }, deviceScaleFactor: size.phone ? 2 : 1, isMobile: size.phone, hasTouch: size.phone })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  let step = 0
  const shot = (name: string) => page.screenshot({ path: resolve(out, `spa-${size.name}-a${String(++step).padStart(2, '0')}-${name}.png`) })
  const at = (what: string) => `[${size.name} autocomplete] ${what}`

  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  await press(page, size, '[data-spa="new"]')
  await press(page, size, '[data-spa-locale="ru-RU"]')
  await press(page, size, '[data-spa="create"]')
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  for (const [cell, text] of [['1:1', '1'], ['2:1', '2'], ['3:1', '3']] as const) await enter(page, size, cell, text)
  await shows(page, '3:1', '3')
  await ready(page)

  // 1. `=СУ`: the list under the field, СУММ first and highlighted.
  const field = size.phone ? '.xs-fx-input' : '.xs-editor'
  if (size.phone) {
    await page.locator('[data-cell="4:1"]').tap()
    await page.locator('.xs-fx').tap()
  } else await page.locator('[data-cell="4:1"]').click()
  await page.keyboard.type('=СУ')
  await page.waitForSelector('[data-xs-functions] [data-xs-function="СУММ"]')
  const list = await page.evaluate(() => [...document.querySelectorAll('[data-xs-function]')].map((li) => li.textContent!.trim()))
  expect(at('the list for =СУ'), list.slice(0, 3), ['СУММ', 'СУММЕСЛИ', 'СУММЕСЛИМН'])
  expect(at('the highlighted item'), await page.evaluate(() => document.querySelector('.xs-fl-on')?.textContent?.trim()), 'СУММ')
  if (size.phone) {
    const heights = await page.evaluate(() => [...document.querySelectorAll<HTMLElement>('[data-xs-function]')].map((li) => li.getBoundingClientRect().height))
    if (heights.some((h) => h < 44)) throw new Error(`${at('items')}: heights ${JSON.stringify(heights)}, want 44 px`)
    await checkNoSideScroll(page, at('the list'))
  }
  await shot('autocomplete')

  // 2. Tab, or a tap on the item: `СУММ(` with the caret after it, the field still open.
  if (size.phone) await page.locator('[data-xs-function="СУММ"]').tap()
  else await page.keyboard.press('Tab')
  await page.waitForSelector('[data-xs-functions]', { state: 'detached' })
  expect(at('the field after the pick'), await page.locator(field).inputValue(), '=СУММ(')
  expect(at('the focus after the pick'), await page.evaluate((sel) => document.activeElement === document.querySelector(sel), field), true)

  // 3. The range and Enter: A4 sums A1:A3.
  await page.keyboard.type('A1:A3)')
  await page.keyboard.press('Enter')
  await shows(page, '4:1', '6')
  await ready(page)
  expect(at('A4 in the book'), cellOf(await bookOf(page), 'A4').formula, 'SUM(A1:A3)')
  await shot('summed')
  if (size.phone) {
    // the formula bar over a formula with references: no target under 44 px
    await page.locator('[data-cell="4:1"]').tap()
    await page.waitForFunction(() => document.querySelector('.xs-fx-text')?.textContent?.includes('СУММ(A1:A3)'))
    await checkTargets(page, at('the formula bar over a formula'))
    await shot('formula-bar')
  }
  console.log(`${at('ok')}: =СУ listed СУММ first, ${size.phone ? 'a tap' : 'Tab'} entered СУММ(, A4 is SUM(A1:A3) = 6`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

/** A file whose A1 and B1 read each other, C1 reads them and E1 does not, with the values Excel would store; opened from the picker. */
async function cycleScenario(browser: Browser, base: string, size: Size): Promise<void> {
  const context = await browser.newContext({ viewport: { width: size.width, height: size.height } })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  const at = (what: string) => `[${size.name} cycle] ${what}`

  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  const cells = { A1: { formula: 'B1/2+1' }, B1: { formula: 'A1' }, C1: { formula: 'A1*10' }, D1: { value: 5 }, E1: { formula: 'D1*2' } }
  const stored = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { value: 2 }, C1: { value: 20 }, E1: { value: 10 } } }] }
  const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells }] }, result: stored })
  const [picker] = await Promise.all([page.waitForEvent('filechooser'), press(page, size, '[data-spa="open"]')])
  await picker.setFiles({ name: 'Цикл.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(data) })
  await shows(page, '1:5', '10')
  await ready(page)
  for (const [cell, text] of [['1:1', '2'], ['1:2', '2'], ['1:3', '20']] as const) await shows(page, cell, text)
  await enter(page, size, '1:4', '7')
  await shows(page, '1:5', '14')
  await ready(page)
  for (const [cell, text] of [['1:1', '2'], ['1:2', '2'], ['1:3', '20']] as const) await shows(page, cell, text)
  await page.screenshot({ path: resolve(out, `spa-${size.name}-c01-cycle.png`) })
  console.log(`${at('ok')}: A1, B1 and C1 hold the file's 2, 2 and 20 through an edit of D1, E1 computed`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

/** A formula's range dragged on the sheet and a reference by an arrow (Point mode), a formula that does not read put back, two typed lines wrapping and fitting their row, its edge dragged and double-clicked. */
async function formulaRowScenario(browser: Browser, base: string, size: Size): Promise<void> {
  const context = await browser.newContext({ viewport: { width: size.width, height: size.height }, permissions: ['clipboard-read', 'clipboard-write'] })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  const at = (what: string) => `[${size.name} formula-row] ${what}`
  const field = () => page.locator('.xs-editor').inputValue()
  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  for (const [cell, text] of [['1:1', '1'], ['2:1', '2']] as const) await enter(page, size, cell, text)
  // =СУММ( and a drag from A1 to A2: the range goes into the field, which stays open
  await page.locator('[data-cell="1:3"]').click()
  await page.keyboard.type('=СУММ(')
  const [a1, a2] = [await middle(page, '[data-cell="1:1"]'), await middle(page, '[data-cell="2:1"]')]
  await page.mouse.move(a1.x, a1.y)
  await page.mouse.down()
  await page.mouse.move(a2.x, a2.y, { steps: 4 })
  await page.mouse.up()
  expect(at('the field after the drag'), await field(), '=СУММ(A1:A2')
  await page.keyboard.type(')')
  await page.keyboard.press('Enter')
  await shows(page, '1:3', '3')
  // = typed over D1 and Left: C1 goes into the field, as Excel's arrows point
  await page.locator('[data-cell="1:4"]').click()
  await page.keyboard.type('=')
  await page.keyboard.press('ArrowLeft')
  expect(at('the field after Left'), await field(), '=C1')
  await page.keyboard.type('*2')
  await page.keyboard.press('Enter')
  await shows(page, '1:4', '6')
  // a formula that does not read, left by a click where no reference may stand: the question, Restore leaves the cell empty
  await page.locator('[data-cell="2:3"]').click()
  await page.keyboard.type('=1)')
  await page.locator('[data-cell="5:5"]').click()
  await page.waitForSelector('[data-xs-alert="formula"]', { timeout: 30_000 })
  await page.screenshot({ path: resolve(out, `spa-${size.name}-p01-formula-question.png`) })
  await page.locator('[data-xs-alert-button="revert"]').click()
  await page.waitForSelector('[data-xs-alert]', { state: 'detached' })
  expect(at('C2 after Restore'), cellOf(await bookOf(page), 'C2').formula, undefined)
  // two lines typed into A3: the cell wraps and its row fits them by itself (xlcalc 0.33), as Excel's
  await page.locator('[data-cell="3:1"]').click()
  await page.keyboard.type('one')
  await page.keyboard.press('Alt+Enter')
  await page.keyboard.type('two')
  await page.keyboard.press('Enter')
  await ready(page)
  type Rows = Record<string, { height?: number; bestFit?: boolean }> | undefined
  const rowOf = async (r: string) => ((await bookOf(page)).sheets[0]!.rows as Rows)?.[r]
  const fitted = await rowOf('3')
  if (!fitted?.bestFit || !(fitted.height! > 15)) throw new Error(`${at('row 3 after two lines')}: ${JSON.stringify(fitted)}, want fitted above 15 pt`)
  expect(at('A3 wraps'), cellOf(await bookOf(page), 'A3').style?.alignment?.wrap, true)
  // its edge dragged 10 px up sets a height by hand, a double click fits the row again
  const grip = async () => (await page.locator('th[data-row="3"] .xs-grip').boundingBox())!
  let g = await grip()
  await page.mouse.move(g.x + g.width / 2, g.y + 3)
  await page.mouse.down()
  await page.mouse.move(g.x + g.width / 2, g.y - 7, { steps: 4 })
  await page.mouse.up()
  await ready(page)
  const byHand = await rowOf('3')
  if (byHand?.bestFit || byHand?.height === fitted.height) throw new Error(`${at('row 3 by hand')}: ${JSON.stringify(byHand)}`)
  g = await grip()
  await page.mouse.dblclick(g.x + g.width / 2, g.y + 3)
  await ready(page)
  expect(at('row 3 fitted again'), await rowOf('3'), fitted)
  // a number wider than its column shows rounded to it; its copy is General without the column, 11 digits, the session's copyText as Excel copies it
  await enter(page, size, '1:5', '1234567,891234')
  await ready(page)
  const shown = (await page.locator('[data-cell="1:5"]').textContent())!.trim()
  await page.locator('[data-cell="1:5"]').click()
  await page.keyboard.press('ControlOrMeta+c')
  // Excel for Mac copies without the last CRLF, Excel for Windows with it
  const want = (await page.evaluate(() => /Mac/.test(navigator.platform))) ? '1234567,891' : '1234567,891\r\n'
  let copied = ''
  for (let i = 0; i < 100 && copied !== want; i++) {
    await page.waitForTimeout(50)
    copied = await page.evaluate(() => navigator.clipboard.readText())
  }
  expect(at('E1 shown'), shown, '1234568')
  expect(at('E1 copied'), copied, want)
  await page.mouse.move(size.width / 2, size.height / 2)
  await page.screenshot({ path: resolve(out, `spa-${size.name}-p02-row-height.png`) })
  console.log(`${at('ok')}: =СУММ(A1:A2) by a drag gives 3, =C1*2 by Left gives 6; =1) asked and put back; two lines in A3 wrap and fit row 3 to ${fitted.height} pt, by hand ${byHand?.height} pt, a double click fits it again; E1 shows ${shown} and copies ${JSON.stringify(copied)}`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

const demo = process.env.DEMO === '1'
const mode = demo ? { mode: 'demo' } : {}
/** A cell copied over a whole column: the app's limit of writes refuses it at once, as a source's refusal, and nothing changes. */
async function limitScenario(browser: Browser, base: string, size: Size): Promise<void> {
  const context = await browser.newContext({ viewport: { width: size.width, height: size.height } })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  const at = (what: string) => `[${size.name} limit] ${what}`
  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)
  for (const [cell, text] of [['1:1', 'x'], ['2:1', '1'], ['3:1', '2']] as const) await enter(page, size, cell, text)
  await ready(page)
  // Control+Down from A1 goes to the end of the block, A3, as Excel's Range.End
  await page.locator('[data-cell="1:1"]').click()
  await page.keyboard.press('Control+ArrowDown')
  await page.waitForFunction(() => document.querySelector('.xs-cell.xs-active')?.getAttribute('data-cell') === '3:1', null, { timeout: 30_000 })
  await page.locator('[data-cell="1:1"]').click()
  await page.keyboard.press('ControlOrMeta+c')
  await page.locator('.xs-head[data-col="2"]').click()
  const t = Date.now()
  await page.keyboard.press('ControlOrMeta+v')
  await page.waitForFunction(() => /100000/.test(document.querySelector('.xs-message')?.textContent ?? ''), null, { timeout: 30_000 })
  const ms = Date.now() - t
  await ready(page)
  expect(at('B1 after the refusal'), cellOf(await bookOf(page), 'B1').value, undefined)
  if (ms > 2000) throw new Error(`${at('refusal')}: ${ms} ms, want under 2 s`)
  await page.screenshot({ path: resolve(out, `spa-${size.name}-l01-limit.png`) })
  console.log(`${at('ok')}: Control+Down from A1 to A3; A1 copied over column B refused in ${ms} ms: ${await page.locator('.xs-message').textContent()}`)
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}

await build({ configFile: config, logLevel: 'error', ...mode })
const server = await preview({ configFile: config, preview: { port: 0 }, logLevel: 'error', ...mode })
const base = server.resolvedUrls!.local[0]!
if (demo) {
  // the demo's paths, its index and no source maps
  if (!base.endsWith('/demo/')) throw new Error(`[demo] served at ${base}, want /demo/`)
  const maps = readdirSync(resolve(root, 'example/dist-demo'), { recursive: true }).filter((f) => String(f).endsWith('.map'))
  if (maps.length) throw new Error(`[demo] source maps in the build: ${maps.join(', ')}`)
  const [index, spa] = await Promise.all([fetch(base).then((r) => r.text()), fetch(`${base}spa.html`).then((r) => r.text())])
  if (index !== spa || !spa.includes('/demo/assets/')) throw new Error('[demo] /demo/ is not the app built under /demo/')
  console.log(`[demo] ok: served at ${base}, /demo/ is the app, no source maps`)
}
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
  await fillCopyScenario(browser, base, desktop)
  await fillCopyScenario(browser, base, phone)
  await autocompleteScenario(browser, base, desktop)
  await autocompleteScenario(browser, base, phone)
  await cycleScenario(browser, base, desktop)
  await limitScenario(browser, base, desktop)
  await formulaRowScenario(browser, base, desktop)
} finally {
  await browser.close()
  await new Promise<void>((done) => server.httpServer.close(() => done()))
}
console.log(`screens in ${out}`)
