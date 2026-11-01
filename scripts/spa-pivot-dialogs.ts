// The dialogs over a PivotTable in the app, for spa-check: a ru-RU workbook's data with dates, a PivotTable inserted from the
// toolbar, then on the desktop the filter menu of Row Labels with Label Filters and its dialog, the cell menu of a value
// with Summarize Values By, Value Field Settings with its tabs and Number Format, Field Settings, PivotTable Options with
// its tabs, Grouping of dates, Insert Calculated Field and the Design tab's Report Layout; on the phone the filter menu
// and Value Field Settings within the screen. Screenshots in test-results/screens/spa-<size>-d*.png.
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
  ['1:3', 'Дата'],
  ['2:1', 'Север'],
  ['2:2', '10'],
  ['2:3', '01.04.2024'],
  ['3:1', 'Юг'],
  ['3:2', '20'],
  ['3:3', '13.05.2024'],
  ['4:1', 'Север'],
  ['4:2', '5'],
  ['4:3', '01.03.2025'],
  ['5:1', 'Восток'],
  ['5:2', '7'],
  ['5:3', '09.03.2025'],
]

const bookOf = (page: Page) => page.evaluate(() => (window as any).__spa.book() as Promise<BookInput>)
async function waitPivot(page: Page, where: string, ok: (t: PivotTable, book: BookInput) => boolean): Promise<void> {
  for (let i = 0; i < 100; i++) {
    const book = await bookOf(page)
    const t = book.sheets.flatMap((s) => s.pivotTables ?? [])[0]
    if (t && ok(t, book)) return
    await page.waitForTimeout(100)
  }
  throw new Error(`${where}: the PivotTable did not change`)
}

export async function pivotDialogsScenario(browser: Browser, base: string, size: Size, out: string): Promise<void> {
  const context = await browser.newContext({ viewport: { width: size.width, height: size.height }, deviceScaleFactor: size.phone ? 2 : 1, isMobile: size.phone, hasTouch: size.phone })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  const at = (what: string) => `[${size.name} pivot dialogs] ${what}`
  let step = 0
  const shot = (name: string) => page.screenshot({ path: resolve(out, `spa-${size.name}-d${String(++step).padStart(2, '0')}-${name}.png`) })
  const press = (selector: string) => (size.phone ? page.locator(selector).first().tap() : page.locator(selector).first().click())
  const ready = () => page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
  const menuAt = async (cell: string) => {
    await page.locator(`[data-cell="${cell}"]`).click({ button: 'right' })
    await page.waitForSelector('[data-xs-menu-box]')
  }
  const noSideScroll = async (what: string) => {
    const { scroll, width } = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, width: window.innerWidth }))
    if (scroll > width) throw new Error(`${at(what)}: a side scroll, ${scroll} > ${width}`)
  }
  const inScreen = async (selector: string, what: string) => {
    const b = await page.locator(selector).first().boundingBox()
    if (!b || b.x < 0 || b.y < 0 || b.x + b.width > size.width + 1 || b.y + b.height > size.height + 1) throw new Error(`${at(what)}: out of the screen, ${JSON.stringify(b)}`)
  }

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
  if (!(await page.locator('[data-tb="pivot-table"]').first().isVisible())) await press('[data-tb="more"]')
  await press('[data-tb="pivot-table"]')
  await page.waitForSelector('[data-xs-pivot-dialog]')
  await press('[data-xs-pivot-dialog] [data-xs-dialog-ok]')
  await page.waitForSelector('[data-xs-pf]', { timeout: 30_000 })
  await press('[data-xs-pf-field="0"] input')
  await waitPivot(page, at('Регион checked'), (t) => t.rows?.[0] === 0)
  await press('[data-xs-pf-field="1"] input')
  await waitPivot(page, at('Сумма checked'), (t) => t.values?.[0]?.field === 1)
  await page.waitForSelector('[data-pivot-button="A3"]', { timeout: 30_000 })

  // the filter menu of Row Labels
  await press('[data-pivot-button="A3"]')
  await page.waitForSelector('[data-pivot-filter-menu]')
  await shot('filter-menu')
  if (size.phone) {
    await inScreen('[data-pivot-filter-menu]', 'the filter menu')
    await noSideScroll('the filter menu')
    await page.keyboard.press('Escape')
    await press('[data-cell="4:2"]')
    // Value Field Settings from the pane's menu of the value field
    await press('[data-xs-pf-area="values"] .xs-pf-chip')
    await press('[data-xs-pf-item="valueFieldSettings"]')
    await page.waitForSelector('[data-pivot-value]')
    await shot('value-settings')
    await inScreen('[data-pivot-value]', 'Value Field Settings')
    await noSideScroll('Value Field Settings')
    await press('[data-pivot-value] [data-xs-dialog-cancel]')
    console.log(`${at('ok')}: the filter menu and Value Field Settings within the screen, no side scroll`)
  } else {
    await page.locator('[data-pfm="label"]').hover()
    await page.waitForSelector('[data-pfm-sub="label"]')
    await shot('label-filters')
    await page.locator('[data-pfm-op="beginsWith"]').click()
    await page.waitForSelector('[data-pivot-filter-dialog]')
    await page.locator('[data-pf-first]').click()
    await page.keyboard.insertText('С')
    await shot('label-filter-dialog')
    await page.locator('[data-pivot-filter-dialog] [data-xs-dialog-ok]').click()
    await waitPivot(page, at('Label Filter'), (t) => t.filters?.[0]?.type === 'captionBeginsWith')
    await page.waitForFunction(() => document.querySelector('[data-cell="4:1"]')?.textContent?.trim() === 'Север', null, { timeout: 30_000 })
    await page.locator('.xs-scroller').focus()
    await page.keyboard.press('Control+z')
    await waitPivot(page, at('Ctrl+Z of the filter'), (t) => !t.filters?.length)

    // the cell menu of a value, Summarize Values By, Value Field Settings with Number Format
    await menuAt('4:2')
    await page.locator('[data-xs-menu="pivotSummarize"]').hover()
    await page.waitForSelector('[data-xs-submenu="pivotSummarize"]')
    await shot('cell-menu-summarize')
    await page.locator('[data-xs-menu="pivotValueSettings"]').click()
    await page.waitForSelector('[data-pivot-value]')
    await shot('value-settings')
    await page.locator('[data-pv-tab="showAs"]').click()
    await page.locator('[data-pv-show-as]').selectOption('difference')
    await shot('value-show-as')
    await page.locator('[data-pv-format]').click()
    await page.waitForSelector('[data-pivot-number-format]')
    await shot('number-format')
    await page.locator('[data-pivot-number-format] [data-xs-dialog-cancel]').click()
    await page.locator('[data-pivot-value] [data-xs-dialog-cancel]').click()

    // Field Settings and PivotTable Options from a label's menu
    await menuAt('4:1')
    await shot('cell-menu-label')
    await page.locator('[data-xs-menu="pivotFieldSettings"]').click()
    await page.waitForSelector('[data-pivot-field]')
    await shot('field-settings')
    await page.locator('[data-pfd-tab="layout"]').click()
    await shot('field-settings-layout')
    await page.locator('[data-pivot-field] [data-xs-dialog-cancel]').click()
    await menuAt('4:1')
    await page.locator('[data-xs-menu="pivotOptions"]').click()
    await page.waitForSelector('[data-pivot-options]')
    for (const tab of ['layout', 'totals', 'display', 'data', 'altText']) {
      await page.locator(`[data-po-tab="${tab}"]`).click()
      await shot(`options-${tab}`)
    }
    await page.locator('[data-pivot-options] [data-xs-dialog-cancel]').click()

    // Grouping of dates: Дата checked onto Rows in place of Регион is grouped by years, quarters and months, as Excel's field list does
    // the period fields named in the workbook's language, as Excel names them in ru-RU
    const named = (book: BookInput, ru: string) => !!book.pivotCaches?.[0]?.fields.some((f) => f.name === `${ru} (Дата)`)
    await page.locator('[data-xs-pf-field="0"] input').click()
    await waitPivot(page, at('Регион unchecked'), (t) => !t.rows?.length)
    await page.locator('[data-xs-pf-field="2"] input').click()
    await waitPivot(page, at('Дата checked'), (t, book) => t.rows?.at(-1) === 2 && named(book, 'Годы') && named(book, 'Кварталы'))
    // Group… on a year opens Grouping over Дата; months and years alone take the quarters away
    await menuAt('4:1')
    await page.locator('[data-xs-menu="pivotGroup"]').click()
    await page.waitForSelector('[data-pivot-group]')
    await page.locator('[data-pg-periods]').selectOption(['months', 'years'])
    await shot('grouping')
    await page.locator('[data-pivot-group] [data-xs-dialog-ok]').click()
    await waitPivot(page, at('Grouping'), (_t, book) => named(book, 'Годы') && !named(book, 'Кварталы'))

    // Fields, Items, & Sets > Calculated Field…, and Design > Report Layout
    await page.locator('[data-cell="4:2"]').click()
    if (!(await page.locator('[data-tb="pivot-calculations"]').first().isVisible())) await page.locator('[data-tb="more"]').click()
    await page.locator('[data-tb="pivot-calculations"]').first().click()
    await page.locator('[data-tb="pivot-calc-field"]').click()
    await page.waitForSelector('[data-pivot-calc]')
    await shot('calculated-field')
    await page.locator('[data-pc-close]').click()
    // the toolbar's Design group may be in its More menu
    if (!(await page.locator('[data-tb="pivot-layout"]').first().isVisible())) await page.locator('[data-tb="more"]').click()
    await page.locator('[data-tb="pivot-layout"]').first().click()
    await page.waitForSelector('[data-tb="pivot-layout-compact"][aria-checked="true"]')
    await shot('design-layout')
    await page.locator('[data-tb="pivot-layout-tabular"]').click()
    await waitPivot(page, at('Report Layout'), (t) => t.layout === 'tabular')
    await shot('tabular')
    console.log(`${at('ok')}: the filter menu with Label Filters and its dialog, Ctrl+Z; the cell menus, Value Field Settings with Number Format, Field Settings, PivotTable Options, Grouping of dates, Calculated Field, Report Layout`)
  }
  if (errors.length) throw new Error(`${at('page errors')}: ${errors.join('; ')}`)
  await context.close()
}
