// The toolbar's Underline, Accounting Number Format and Orientation menus in the app (spa.html, the local source): each menu
// open in English and in Russian, with rows 1 and 2, then the keyboard on a new en-US workbook: Space and Enter on a split
// button's part and arrow and on an item, Esc closing a menu, each command's effect read from the workbook. Screenshots go to
// test-results/screens/ with the others; called from screens.ts on the desktop.
import { resolve } from 'node:path'
import type { Page } from 'playwright'
import type { BookInput, CellStyle } from 'xlcalc'
import { parseCell, parseRange } from 'xlcalc/a1'

interface Size {
  name: string
  mobile: boolean
}
const MENUS = ['underline-menu', 'accounting-menu', 'orientation']

async function ready(page: Page): Promise<void> {
  await page.waitForSelector('[data-cell="1:1"]')
  await page.waitForFunction(() => ['Ready', 'Готово'].includes(document.querySelector('[data-spa-status]')?.textContent?.trim() ?? ''), null, { timeout: 30_000 })
}
/** A cell of the first sheet as the workbook holds it: its own entry, otherwise the last style layer over it (an empty cell's formatting). */
async function cellOf(page: Page, addr: string): Promise<{ style?: CellStyle; numberFormat?: unknown }> {
  const book = (await page.evaluate(() => (window as any).__spa.book())) as BookInput
  const sheet = book.sheets[0]!
  const at = parseCell(addr)!
  const layer = [...(sheet.looks ?? [])].reverse().find((l) => {
    const r = parseRange(l.range)
    return r && at.row >= r.top && at.row <= r.bottom && at.col >= r.left && at.col <= r.right
  })
  const c = ((sheet.cells ?? {}) as Record<string, { style?: number; numberFormat?: unknown }>)[addr] ?? layer ?? {}
  return { style: c.style === undefined ? undefined : book.styles?.[c.style], numberFormat: c.numberFormat }
}
function expect(what: string, got: unknown, want: unknown): void {
  if (JSON.stringify(got) !== JSON.stringify(want)) throw new Error(`ribbon menus: ${what}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`)
}
const open = (page: Page) => page.locator('[role="menu"]').count()

export async function ribbonMenuScreens(page: Page, base: string, size: Size, out: string): Promise<void> {
  for (const lang of ['en', 'ru']) {
    for (const rows of [1, 2]) {
      await page.goto(`${base}spa.html?lang=${lang}&rows=${rows}`)
      await ready(page)
      for (const arrow of MENUS) {
        // a group that does not fit is in the menu under More
        if (await page.locator(`[data-tb-menu] [data-tb="${arrow}"]`).count()) await page.locator('[data-tb="more"]').click()
        await page.locator(`[data-tb="${arrow}"]`).click()
        await page.waitForSelector('[role="menu"] [data-tb]')
        await page.mouse.move(5, size.mobile ? 800 : 1150)
        await page.screenshot({ path: resolve(out, `${size.name}-71-ribbon-${lang}-rows${rows}-${arrow}.png`) })
        await page.keyboard.press('Escape')
        expect(`${lang} rows ${rows}: Esc closes ${arrow}'s menu`, await open(page), 0)
      }
    }
  }
  console.log(`ribbon menus ${size.name}: Underline, Accounting Number Format and Orientation open in en and ru, rows 1 and 2, closed by Esc`)

  // The keyboard on the English app's en-US workbook, A1 selected
  await page.goto(`${base}spa.html?lang=en&rows=2`)
  await ready(page)
  await page.locator('[data-cell="1:1"]').click()
  const key = async (hook: string, k: 'Enter' | 'Space') => {
    await page.locator(`[data-tb="${hook}"]`).focus()
    await page.keyboard.press(k)
  }
  const settled = (what: () => Promise<unknown>, want: unknown, label: string) => waitFor(async () => JSON.stringify(await what()) === JSON.stringify(want), label, page)
  await key('underline', 'Space')
  await settled(async () => (await cellOf(page, 'A1')).style?.font?.underline, 'single', 'Space on Underline underlines A1')
  await key('underline-menu', 'Enter')
  expect('Enter on the arrow opens Underline’s menu', await open(page), 1)
  await key('underline-double', 'Enter')
  expect('Enter on an item closes the menu', await open(page), 0)
  await settled(async () => (await cellOf(page, 'A1')).style?.font?.underline, 'double', 'Enter on Double Underline')
  await page.waitForSelector('[data-tb="underline"][aria-pressed="true"][data-tip="Double Underline"]')
  await key('accounting', 'Enter')
  await settled(async () => (await cellOf(page, 'A1')).numberFormat, 44, 'Enter on Accounting Number Format')
  await key('accounting-menu', 'Space')
  expect('Space on the arrow opens the currencies', await page.locator('[role="menu"] [data-tb^="accounting-"]').first().getAttribute('data-tb'), 'accounting-usd')
  await page.keyboard.press('Escape')
  expect('Esc closes the currencies', await open(page), 0)
  await key('orientation', 'Enter')
  await key('orientation-up', 'Space')
  await settled(async () => (await cellOf(page, 'A1')).style?.alignment?.rotation, 90, 'Space on Rotate Text Up')
  await key('orientation', 'Space')
  await page.waitForSelector('[data-tb="orientation-up"][aria-checked="true"]')
  await page.screenshot({ path: resolve(out, `${size.name}-72-ribbon-orientation-checked.png`) })
  await page.keyboard.press('Escape')
  expect('Esc closes Orientation', await open(page), 0)
  console.log(`ribbon menus ${size.name}: Space and Enter on Underline's part, arrow and item, on Accounting's part and arrow, on Orientation and its item; Esc closes`)
}

async function waitFor(ok: () => Promise<boolean>, label: string, page: Page): Promise<void> {
  const end = Date.now() + 15_000
  while (!(await ok())) {
    if (Date.now() > end) throw new Error(`ribbon menus: ${label}: not reached`)
    await page.waitForTimeout(100)
  }
}
