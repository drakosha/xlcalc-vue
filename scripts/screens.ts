// Screenshots of the example at a phone's and a desktop's width: every sheet of the "all formatting" book and the editing states.
// Run: npm run screens (node 22.18+ strips the types). Output: test-results/screens/, not in git.
import { mkdirSync } from 'node:fs'
import { resolve } from 'node:path'
import { chromium, type Page } from 'playwright'
import { createServer } from 'vite'

const root = resolve(import.meta.dirname, '..')
const out = resolve(root, 'test-results/screens')
mkdirSync(out, { recursive: true })

const SHEETS = ['Оформление', 'Форматы', 'Гистограммы', 'Ссылки', 'Примечания', 'Формулы', 'Объединения', 'Размеры', 'Закрепление', 'Таблица', 'Длинный лист', 'Запертые', 'Перелив', 'Плашка']
const SIZES = [
  { name: '390', width: 390, height: 844, mobile: true },
  { name: '1920', width: 1920, height: 1200, mobile: false },
]
const quote = (name: string) => (/^[\p{L}\p{N}_]+$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`)

const server = await createServer({ configFile: resolve(root, 'vite.example.config.ts'), server: { port: 0 }, logLevel: 'error' })
await server.listen()
const base = server.resolvedUrls!.local[0]!
const browser = await chromium.launch()

async function settle(page: Page): Promise<void> {
  await page.waitForSelector('.xs-table')
  await page.waitForFunction(() => !document.querySelector('.xs-loading'))
  await page.waitForTimeout(250)
}
async function shot(page: Page, size: string, name: string): Promise<void> {
  await page.screenshot({ path: resolve(out, `${size}-${name}.png`) })
}
async function open(page: Page, place: string, extra = ''): Promise<void> {
  await page.goto(`${base}?sheet=${encodeURIComponent(place)}${extra}`)
  await settle(page)
}

for (const size of SIZES) {
  const context = await browser.newContext({
    viewport: { width: size.width, height: size.height },
    deviceScaleFactor: size.mobile ? 2 : 1,
    isMobile: size.mobile,
    hasTouch: size.mobile,
  })
  const page = await context.newPage()
  page.on('pageerror', (e) => console.error(`[${size.name}] page error:`, e.message))
  for (const [i, name] of SHEETS.entries()) {
    await open(page, quote(name))
    await shot(page, size.name, `${String(i + 1).padStart(2, '0')}-${name}`)
  }

  // A formula's ranges highlighted, its references in the bar.
  await open(page, 'Формулы!B6')
  await shot(page, size.name, '20-formula-refs')
  // A selection with its totals.
  await open(page, 'Форматы!B2:C11')
  await shot(page, size.name, '21-selection-totals')
  // A jump far down a long sheet flashes the place.
  await open(page, "'Длинный лист'!B30000")
  await page.waitForTimeout(300)
  await shot(page, size.name, '22-jump-flash')

  // Editing in place: typing over a cell, then the edit in flight.
  await open(page, 'Формулы!C2')
  if (size.mobile) {
    await page.locator('.xs-fx').tap()
    await page.keyboard.type('=A2*2')
    await shot(page, size.name, '23-edit-bar')
  } else {
    await page.locator('.xs-scroller').focus()
    await page.keyboard.type('=A2*2')
    await shot(page, size.name, '23-edit-cell')
  }
  await page.keyboard.press('Enter')
  await page.waitForTimeout(30)
  await shot(page, size.name, '24-edit-pending')
  await page.waitForTimeout(400)
  await shot(page, size.name, '25-edit-done')

  // A conflict keeps the text in the field and explains.
  await open(page, 'Формулы!C3', '&fail=conflict')
  if (size.mobile) await page.locator('.xs-fx').tap()
  else await page.locator('.xs-scroller').focus()
  await page.keyboard.type('77')
  await page.keyboard.press('Enter')
  await page.waitForTimeout(700)
  await shot(page, size.name, '26-conflict')

  // A locked cell is explained, not sent.
  await open(page, 'Запертые!A3')
  if (size.mobile) await page.locator('.xs-fx').tap()
  else {
    await page.locator('.xs-scroller').focus()
    await page.keyboard.type('5')
  }
  await page.waitForTimeout(100)
  await shot(page, size.name, '27-locked')
  await context.close()
}

await browser.close()
await server.close()
console.log(`screens in ${out}`)
