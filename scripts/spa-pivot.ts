// The app's PivotTables in a browser, for spa-check: a ru-RU workbook's data, Insert > PivotTable from the toolbar with the
// region it proposes, the PivotTable Fields pane on the new sheet, two fields checked (text onto Rows, numbers onto
// Values), on the desktop the Rows field dragged with the mouse onto Columns and Ctrl+Z, on the phone the pane below the
// grid without a horizontal scroll. Screenshots in test-results/screens/spa-<size>-v*.png.
import { resolve } from 'node:path'
import type { Browser, Page } from 'playwright'
import type { BookInput, PivotTable } from 'xlcalc'

interface Size {
  name: string
  width: number
  height: number
  phone: boolean
}

const DATA: [string, string][] = [
  ['1:1', 'Регион'],
  ['1:2', 'Сумма'],
  ['2:1', 'Север'],
  ['2:2', '10'],
  ['3:1', 'Юг'],
  ['3:2', '20'],
  ['4:1', 'Север'],
  ['4:2', '5'],
]

const pivotOf = async (page: Page): Promise<PivotTable | undefined> => {
  const book = await page.evaluate(() => (window as any).__spa.book() as Promise<BookInput>)
  return book.sheets.flatMap((s) => s.pivotTables ?? [])[0]
}
async function waitPivot(page: Page, where: string, ok: (t: PivotTable) => boolean): Promise<PivotTable> {
  for (let i = 0; i < 100; i++) {
    const t = await pivotOf(page)
    if (t && ok(t)) return t
    await page.waitForTimeout(100)
  }
  const t = await pivotOf(page)
  throw new Error(`${where}: the PivotTable is ${JSON.stringify(t && { rows: t.rows, columns: t.columns, values: t.values })}`)
}

export async function pivotScenario(browser: Browser, base: string, size: Size, out: string): Promise<void> {
  const context = await browser.newContext({ viewport: { width: size.width, height: size.height }, deviceScaleFactor: size.phone ? 2 : 1, isMobile: size.phone, hasTouch: size.phone })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  const at = (what: string) => `[${size.name} pivot] ${what}`
  let step = 0
  const shot = (name: string) => page.screenshot({ path: resolve(out, `spa-${size.name}-v${String(++step).padStart(2, '0')}-${name}.png`) })
  const press = (selector: string) => (size.phone ? page.locator(selector).first().tap() : page.locator(selector).first().click())
  const ready = () => page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready()
  for (const [cell, text] of DATA) {
    if (size.phone) {
      await page.locator(`[data-cell="${cell}"]`).tap()
      await page.locator('.xs-fx').tap()
      await page.keyboard.press('ControlOrMeta+A')
      await page.keyboard.insertText(text)
    } else {
      await page.locator(`[data-cell="${cell}"]`).click()
      // a Cyrillic first letter typed through insertText has no keydown: F2 opens the edit first
      await page.keyboard.press('F2')
      await page.keyboard.insertText(text)
    }
    await page.keyboard.press('Enter')
  }
  await ready()
  await press('[data-cell="2:1"]')
  // the toolbar's button, in its More menu on a phone
  if (!(await page.locator('[data-tb="pivot-table"]').first().isVisible())) await press('[data-tb="more"]')
  await press('[data-tb="pivot-table"]')
  await page.waitForSelector('[data-xs-pivot-dialog]')
  const proposed = await page.locator('[data-xs-pivot-source]').inputValue()
  if (proposed !== 'Лист1!$A$1:$B$4') throw new Error(`${at('Table/Range')}: ${proposed}`)
  await shot('insert-dialog')
  await press('[data-xs-pivot-dialog] [data-xs-dialog-ok]')
  await page.waitForSelector('[data-xs-pf]', { timeout: 30_000 })
  const tabs = await page.locator('[role="tab"]').allTextContents()
  if (tabs[0]?.trim() !== 'Лист2') throw new Error(`${at('new sheet')}: the tabs are ${tabs.join(', ')}`)
  await press('[data-xs-pf-field="0"] input')
  await waitPivot(page, at('Region checked'), (t) => t.rows?.[0] === 0)
  await press('[data-xs-pf-field="1"] input')
  await waitPivot(page, at('Amount checked'), (t) => t.values?.[0]?.field === 1)
  await page.waitForFunction(() => document.querySelector('[data-cell="4:2"]')?.textContent?.trim() === '15', null, { timeout: 30_000 })
  await shot('pane')
  if (size.phone) {
    const { scroll, width } = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, width: window.innerWidth }))
    if (scroll > width) throw new Error(`${at('side scroll')}: ${scroll} > ${width}`)
    const pane = await page.locator('[data-xs-pf]').boundingBox()
    const grid = await page.locator('.xs-body').boundingBox()
    if (!pane || !grid || pane.y < grid.y + grid.height - 1) throw new Error(`${at('placement')}: the pane is not below the grid`)
    console.log(`${at('ok')}: Insert PivotTable from the More menu proposed Лист1!$A$1:$B$4, the pane below the grid on Лист2, Регион onto Rows and Сумма onto Values`)
  } else {
    // a drag with the mouse: the Rows field onto Columns
    const from = (await page.locator('[data-xs-pf-area="rows"] .xs-pf-chip').boundingBox())!
    const to = (await page.locator('[data-xs-pf-area="columns"] .xs-pf-chips').boundingBox())!
    await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2)
    await page.mouse.down()
    await page.mouse.move(from.x + from.width / 2 + 10, from.y + from.height / 2 - 10, { steps: 3 })
    await page.mouse.move(to.x + to.width / 2, to.y + to.height / 2, { steps: 8 })
    await shot('dragging')
    await page.mouse.up()
    await waitPivot(page, at('dragged onto Columns'), (t) => t.columns?.[0] === 0 && !t.rows?.length)
    await shot('dragged')
    await page.locator('.xs-scroller').focus()
    await page.keyboard.press('Control+z')
    await waitPivot(page, at('Ctrl+Z'), (t) => t.rows?.[0] === 0 && !t.columns?.length)
    console.log(`${at('ok')}: Insert PivotTable proposed Лист1!$A$1:$B$4 and put it on Лист2 at A3 with the pane, Регион onto Rows and Сумма onto Values, Регион dragged onto Columns, Ctrl+Z back`)
  }
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}
