// The browser checks of Excel's zoom and Watch Window on the example: hit testing at 50% and 200% (a click, a range
// pointed in a formula, a column edge, the editor over its cell), Ctrl+wheel's steps, a pinch on a phone, and the Watch Window's
// values following an edit. Screenshots go to test-results/screens/ with the others; called from screens.ts.
import { resolve } from 'node:path'
import type { Page } from 'playwright'
import { pinch } from './touch.ts'

interface Size {
  name: string
  mobile: boolean
}

const FROZEN = encodeURIComponent('Закрепление!A1')

async function ready(page: Page): Promise<void> {
  await page.waitForSelector('.xs-table')
  await page.waitForFunction(() => !document.querySelector('.xs-loading'))
  await page.waitForTimeout(250)
}
const state = (page: Page) => page.textContent('[data-state]')
const rectOf = (page: Page, cell: string) => page.evaluate((c) => document.querySelector(`[data-cell="${c}"]`)!.getBoundingClientRect().toJSON() as DOMRect, cell)
function expect(what: string, got: unknown, want: unknown): void {
  if (JSON.stringify(got) !== JSON.stringify(want)) throw new Error(`zoom: ${what}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`)
}
const near = (a: number, b: number, tol = 2) => Math.abs(a - b) <= tol

/** A click, a pointed range, the editor and a column edge on the frozen sheet scrolled down and right, at a zoom. */
async function hitTests(page: Page, base: string, size: Size, out: string, zoom: number): Promise<void> {
  await page.goto(`${base}?sheet=${FROZEN}&zoom=${zoom}&lang=en`)
  await ready(page)
  // scrolled by 40 rows and 6 columns of the unzoomed grid, so the frozen panes and the body differ
  await page.evaluate((z) => {
    const el = document.querySelector<HTMLElement>('.xs-scroller')!
    el.scrollTo({ top: 800 * z, left: 430 * z })
  }, zoom / 100)
  await ready(page)
  // a body cell of row 50 in view, right of the frozen column (cells are drawn beyond the viewport and under the pane too), clicked in its visible part
  const body = await page.evaluate(() => {
    const view = document.querySelector('.xs-scroller')!.getBoundingClientRect()
    const pane = document.querySelector('td[data-cell="50:1"]')!.getBoundingClientRect()
    const tds = [...document.querySelectorAll<HTMLElement>('td[data-cell^="50:"]')].map((td) => ({ td, r: td.getBoundingClientRect() }))
    // wholly in view where one is, otherwise the visible part of one (a narrow viewport zoomed in)
    const whole = tds.find(({ r }) => r.left >= pane.right - 0.5 && r.right <= view.right - 2)
    for (const { td, r } of whole ? [whole] : tds) {
      const from = Math.max(r.left, pane.right)
      const to = Math.min(r.right, view.right - 2)
      if (to - from > 8) return { cell: td.dataset.cell!, x: (from + to) / 2, y: r.top + r.height * 0.75 }
    }
    return null
  })
  if (!body) throw new Error(`zoom: ${zoom}%: no body cell in view`)
  const r = await rectOf(page, body.cell)
  await page.mouse.click(body.x, body.y)
  await page.waitForTimeout(80)
  const [row, col] = body.cell.split(':').map(Number) as [number, number]
  const letters = (c: number) => String.fromCharCode(64 + c)
  expect(`a click at ${zoom}%`, await state(page), `Закрепление!${letters(col)}${row}`)
  // a frozen cell in column A at the same height
  const a = await rectOf(page, `${row}:1`)
  await page.mouse.click(a.x + 3, a.y + a.height / 2)
  await page.waitForTimeout(80)
  expect(`a click on a frozen cell at ${zoom}%`, (await state(page))?.split(' ')[0], `Закрепление!A${row}`)
  if (size.mobile) return
  // Point mode: a drag from the cell to the one four rows below and one column right writes their range into a formula
  // (hit testing by the geometry under the zoom, not by the element under the pointer)
  const from = await rectOf(page, `${row - 2}:${col}`)
  const to = await rectOf(page, `${row + 2}:${col + 1}`)
  await page.mouse.click(body.x, body.y)
  await page.keyboard.type('=')
  await page.mouse.move(from.x + from.width * 0.9, from.y + from.height * 0.9)
  await page.mouse.down()
  await page.mouse.move(to.x + to.width * 0.1, to.y + to.height * 0.1, { steps: 6 })
  await page.mouse.up()
  const typed = await page.inputValue('.xs-editor')
  expect(`a range pointed at ${zoom}%`, typed, `=${letters(col)}${row - 2}:${letters(col + 1)}${row + 2}`)
  await page.keyboard.press('Escape')
  // the editor lies over its cell
  await page.mouse.dblclick(body.x, body.y)
  await page.waitForTimeout(80)
  const ed = (await page.locator('.xs-editor:not(.xs-editor-idle)').boundingBox())!
  if (!near(ed.x, r.x) || !near(ed.y, r.y) || ed.height < r.height - 1) throw new Error(`zoom: the editor at ${zoom}% is at ${JSON.stringify(ed)}, its cell at ${JSON.stringify(r)}`)
  await page.keyboard.press('Escape')
  // a column edge dragged by 40 page pixels widens the column by 40 / zoom unzoomed pixels
  const head = (await page.locator(`th.xs-head[data-col="${col}"]`).boundingBox())!
  const widthOf = (c: number) => page.evaluate((c) => parseFloat(getComputedStyle(document.querySelector(`th.xs-head[data-col="${c}"]`)!).width), c)
  const before = await widthOf(col)
  await page.mouse.move(head.x + head.width - 1, head.y + head.height / 2)
  await page.mouse.down()
  await page.mouse.move(head.x + head.width + 39, head.y + head.height / 2, { steps: 4 })
  await page.mouse.up()
  const grew = (await widthOf(col)) - before
  if (!near(grew, 40 / (zoom / 100), 1.5)) throw new Error(`zoom: a column edge dragged by 40 px at ${zoom}% grew the column by ${grew} px`)
  await page.screenshot({ path: resolve(out, `${size.name}-70-zoom-${zoom}.png`) })
}

/** Ctrl+wheel's notches step by 15%, and the status bar's percentage reads them. */
async function wheel(page: Page, base: string): Promise<void> {
  await page.goto(`${base}?sheet=${FROZEN}&lang=en`)
  await ready(page)
  const box = (await page.locator('.xs-scroller').boundingBox())!
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
  await page.keyboard.down('Control')
  await page.mouse.wheel(0, -100)
  await page.waitForTimeout(80)
  expect('Ctrl+wheel up', await page.textContent('[data-xs-zoom-level]'), '115%')
  await page.mouse.wheel(0, 100)
  await page.mouse.wheel(0, 100)
  await page.waitForTimeout(80)
  await page.keyboard.up('Control')
  expect('Ctrl+wheel down twice', await page.textContent('[data-xs-zoom-level]'), '85%')
}

/** Two fingers spread on a phone zoom the sheet in, about their center. */
async function pinchTest(page: Page, base: string, out: string, size: Size): Promise<void> {
  await page.goto(`${base}?sheet=${FROZEN}&lang=en`)
  await ready(page)
  const box = (await page.locator('.xs-scroller').boundingBox())!
  await pinch(page, box.x + box.width / 2, box.y + box.height / 2, 80, 160)
  await page.waitForTimeout(150)
  const zoom = await page.evaluate(() => Number(getComputedStyle(document.querySelector('.xs-zoom')!).zoom))
  if (!(zoom > 1.6 && zoom < 2.4)) throw new Error(`zoom: a pinch from 80 to 160 px gave ${zoom}`)
  await page.screenshot({ path: resolve(out, `${size.name}-71-pinch.png`) })
}

/** The Watch Window: Add Watch for A2:B3 (values, a formula), a value following an edit, a double click going to a cell from another sheet. */
async function watchWindow(page: Page, base: string, out: string, size: Size): Promise<void> {
  await page.goto(`${base}?sheet=${encodeURIComponent('Формулы!A2:B3')}&lang=en`)
  await ready(page)
  await page.evaluate(() => document.querySelector<HTMLElement>('[data-watch-window]')!.click())
  await page.waitForSelector('[data-xs-watch]')
  await page.locator('[data-xs-watch-add]').click()
  await page.waitForTimeout(400)
  const rows = await page.$$eval('[data-xs-watch-row]', (trs) => trs.map((tr) => [...tr.querySelectorAll('td')].map((td) => td.textContent!.trim())))
  expect('watched cells', rows.map((r) => r[3]), ['A2', 'B2', 'A3', 'B3'])
  const formula = await page.$eval('[data-xs-watch-row="Формулы!B2"] [data-xs-watch-formula]', (td) => td.textContent!.trim())
  if (!formula.startsWith('=')) throw new Error(`zoom: the watched B2 shows no formula: ${formula}`)
  await page.screenshot({ path: resolve(out, `${size.name}-72-watch.png`) })
  // another tab's edit of A2 reaches the Watch Window
  await page.evaluate(() => [...document.querySelectorAll<HTMLButtonElement>('.top button')].find((b) => b.textContent!.includes('another tab'))!.click())
  await page.waitForTimeout(500)
  const a2 = await page.$eval('[data-xs-watch-row="Формулы!A2"] [data-xs-watch-value]', (td) => td.textContent!.trim())
  const shown = await page.$eval('[data-cell="2:1"]', (td) => td.textContent!.trim())
  expect('the watched value after an edit elsewhere', a2, shown)
  if (size.mobile) return
  await page.locator('.xs-tab', { hasText: 'Оформление' }).click()
  await ready(page)
  await page.locator('[data-xs-watch-row="Формулы!A3"]').dblclick()
  await page.waitForTimeout(400)
  expect('a double click on a watch', (await state(page))?.split(' ')[0], 'Формулы!A3')
}

export async function zoomScreens(page: Page, base: string, size: Size, out: string): Promise<void> {
  for (const zoom of [50, 100, 200]) await hitTests(page, base, size, out, zoom)
  if (size.mobile) await pinchTest(page, base, out, size)
  else await wheel(page, base)
  await watchWindow(page, base, out, size)
  console.log(`zoom and watch ${size.name}: checked`)
}
