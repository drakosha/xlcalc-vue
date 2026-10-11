// Screenshots of the example at a phone's and a desktop's width: every sheet of the "all formatting" book and the editing states.
// Run: npm run screens (node 22.18+ strips the types). Output: test-results/screens/, not in git.
import { mkdirSync } from 'node:fs'
import { resolve } from 'node:path'
import { chromium, type Page } from 'playwright'
import { createServer } from 'vite'
import { quoteSheetIfNeeded } from 'xlcalc/a1'

const root = resolve(import.meta.dirname, '..')
const out = resolve(root, 'test-results/screens')
mkdirSync(out, { recursive: true })

const SHEETS = ['Оформление', 'Форматы', 'Гистограммы', 'Ссылки', 'Примечания', 'Формулы', 'Объединения', 'Размеры', 'Закрепление', 'Таблица', 'Длинный лист', 'Запертые', 'Перелив', 'Плашка', 'Фильтр']
const SIZES = [
  { name: '390', width: 390, height: 844, mobile: true },
  { name: '1920', width: 1920, height: 1200, mobile: false },
]

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

/** Underline, double underline and strike must be computed on the element carrying the text (they do not pass to an overflow run). */
async function checkDecorations(page: Page): Promise<void> {
  await open(page, 'Оформление!A1')
  const got = await page.evaluate(() =>
    ['10:1', '11:1', '12:1'].map((cell) => {
      const td = document.querySelector(`[data-cell="${cell}"]`)!
      const el = [...td.querySelectorAll('.xs-in *')].find((e) => e.children.length === 0 && e.textContent!.trim() !== '') as HTMLElement
      const cs = getComputedStyle(el)
      return `${cell} ${el.textContent!.trim()}: ${cs.textDecorationLine} ${cs.textDecorationStyle}`
    }),
  )
  const want = ['underline solid', 'underline double', 'line-through solid']
  const bad = got.filter((g, i) => !g.endsWith(want[i]!))
  if (bad.length) throw new Error(`text decoration is missing on the text: ${bad.join('; ')}`)
  console.log(`decorations: ${got.join('; ')}`)
}

/** Every button, field and tab outside the grid at least 44 × 44 px on a phone (the address box by its height); a desktop unchanged. */
async function checkTargets(page: Page, mobile: boolean): Promise<void> {
  const boxes = await page.evaluate(() =>
    [...document.querySelectorAll<HTMLElement>('.xs-sheet button, .xs-sheet input, .xs-sheet textarea, .xs-sheet [role=tab], .xs-sheet a')]
      .filter((el) => !el.closest('.xs-table') && !el.matches('.xs-editor-idle, .xs-fx-idle') && el.getClientRects().length > 0)
      .map((el) => {
        const target = el.matches('input[type=checkbox]') ? (el.closest('label') ?? el) : el
        const r = target.getBoundingClientRect()
        return { what: `${el.tagName.toLowerCase()}.${el.className || el.getAttribute('title') || ''}`.slice(0, 60), w: Math.round(r.width), h: Math.round(r.height) }
      }),
  )
  if (mobile) {
    const small = boxes.filter((b) => b.h < 44 || (b.w < 44 && !b.what.includes('xs-address')))
    if (small.length) throw new Error(`touch targets under 44 px: ${JSON.stringify(small)}`)
  } else {
    const tool = boxes.find((b) => b.what.startsWith('button.xs-tool'))
    if (!tool || tool.w !== 24 || tool.h !== 22) throw new Error(`a desktop tool changed: ${JSON.stringify(tool)}`)
  }
  return void console.log(`targets ${mobile ? '390' : '1920'}: ${boxes.length} checked`)
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
  await checkDecorations(page)
  if (process.env.SCREENS_CHECK_ONLY) {
    await context.close()
    continue
  }
  for (const [i, name] of SHEETS.entries()) {
    await open(page, quoteSheetIfNeeded(name))
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

  // Frozen panes stay while the body scrolls; a whole column of a long sheet counts its cells only.
  await open(page, 'Закрепление!M150')
  await page.waitForTimeout(2200)
  await shot(page, size.name, '28-freeze-scrolled')
  await open(page, "'Длинный лист'!C:C")
  await page.waitForTimeout(2200)
  await shot(page, size.name, '29-column-count')

  // Targets for a finger on a phone, with the search bar and the funnel's list open; a desktop's stay as they were.
  await open(page, 'Фильтр!A1')
  await page.waitForTimeout(2200)
  await checkTargets(page, size.mobile)
  await page.locator('[data-xs-find]').click()
  await checkTargets(page, size.mobile)
  await page.locator('[data-xs-find]').click()
  if (size.mobile) {
    await page.locator('.xs-bar [data-xs-funnel]').click()
    await page.waitForSelector('[data-xs-values] .xs-value-all')
    await checkTargets(page, true)
    await page.locator('.xs-popover-head button').last().click()
  }

  // 0.2.0: search with its hits, the funnel of a column (a phone: the whole screen), a filtered and sorted sheet, a recompute.
  await open(page, "'Длинный лист'!A1")
  await page.locator('[data-xs-find]').click()
  await page.locator('.xs-find-input').fill('K-0300')
  await page.waitForSelector('.xs-cell.xs-hit-now')
  await page.waitForTimeout(400)
  await shot(page, size.name, '30-find')
  await open(page, 'Фильтр!A1')
  // A frozen cell that flashes after a jump must stay sticky in its pane.
  const position = await page.evaluate(() => getComputedStyle(document.querySelector('[data-cell="1:1"]')!).position)
  if (position !== 'sticky') throw new Error(`a flashing frozen cell is ${position}, not sticky`)
  const funnel = async (col: number) => {
    if (!size.mobile) return page.locator('.xs-head').nth(col - 1).locator('.xs-funnel').click({ force: true })
    await page.locator(`.xs-cell[data-cell$=":${col}"]`).nth(1).tap()
    await page.locator('.xs-bar [data-xs-funnel]').click()
  }
  await funnel(2)
  await page.waitForSelector('[data-xs-values] .xs-value-all')
  await shot(page, size.name, '31-funnel')
  await page.locator('.xs-value', { hasText: 'отклонён' }).click()
  await page.locator('.xs-value', { hasText: '(пусто)' }).click()
  await page.locator('[data-xs-apply]').click()
  await page.waitForTimeout(300)
  await funnel(3)
  await page.waitForSelector('[data-xs-values] .xs-value-all')
  await page.locator('.xs-popover button[title="Сортировать по убыванию"]').click()
  await page.waitForSelector('[data-xs-filter-status]')
  await page.waitForFunction(() => !document.querySelector('.xs-loading'))
  await page.waitForTimeout(300)
  await shot(page, size.name, '32-filtered')
  await page.locator('[data-xs-filter-clear]').click()
  await page.waitForTimeout(200)
  await page.goto(`${base}?sheet=${encodeURIComponent('Форматы')}&recompute=4000`)
  await page.waitForSelector('[data-xs-wait]')
  await page.waitForTimeout(1600)
  await shot(page, size.name, '33-recompute')

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
