// The local source in a browser: the example's local.html built for production (the pool's thread bundled by Vite) and
// served, then the dev server once. Measures the opening of the synthetic book and edits, checks a Web Worker computes,
// takes screenshots at 390 × 844 and 1920 × 1200. Run: npm run check:local. Output: test-results/screens/local-*.png.
import { mkdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { chromium, type Browser, type Page } from 'playwright'
import { build, createServer, preview } from 'vite'
import { registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { writeXlsx } from 'xlcalc/xlsx'

const root = resolve(import.meta.dirname, '..')
const config = resolve(root, 'vite.example.config.ts')
const out = resolve(root, 'test-results/screens')
mkdirSync(out, { recursive: true })
const ROWS = Number(process.env.ROWS ?? 3400)
const RUNS = 3

interface Opened {
  drawn: number
  ready: number
  built: number
  workers: number
}

const median = (xs: number[]) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)]!
const fmt = (xs: number[]) => `median ${Math.round(median(xs))} ms (${xs.map((x) => Math.round(x)).join(', ')})`

/** Open the page at a place; the time from navigation to the first drawn cell of the place's sheet. */
async function open(page: Page, base: string, place: string, cell: string): Promise<Opened> {
  await page.goto(`${base}local.html?rows=${ROWS}&sheet=${encodeURIComponent(place)}`)
  await page.waitForFunction((c) => !!document.querySelector(`[data-cell="${c}"]`)?.textContent?.trim(), cell, { timeout: 60_000, polling: 'raf' })
  const drawn = await page.evaluate(() => performance.now())
  await page.waitForFunction(() => (window as any).__xs?.timings.ready > 0)
  const t = await page.evaluate(() => (window as any).__xs.timings as { built: number; ready: number })
  return { drawn, ready: t.ready, built: t.built, workers: page.workers().length }
}

/** What an edit showed, in ms from Enter: the watched cell's new text, the first cell of `looks` recoloured (null: none within the watch), frames drawn with `steady` empty. */
interface Watched {
  shown: number
  look: number | null
  blank: number
}

/** How long frames are watched after Enter: a recolour and an empty frame come well within it. */
const WATCH_MS = 1500

/**
 * Type into the active cell (the formula bar on a phone) and press Enter. Every frame until WATCH_MS after Enter is looked
 * at: when `watch` shows other text, when a drawn cell of `looks` changes its background (a conditional format's repaint),
 * and whether `steady`, a cell that always has text, was drawn empty (the windows dropped and read again).
 */
async function typeAndWatch(page: Page, text: string, watch: string, phone: boolean, looks: string[] = [], steady = '3:1'): Promise<Watched> {
  if (phone) {
    // the bar opens with the cell's text, as Excel's does: replace it
    await page.locator('.xs-fx').tap()
    await page.keyboard.press('ControlOrMeta+A')
  } else await page.locator('.xs-scroller').focus()
  await page.keyboard.type(text)
  await page.evaluate(
    ([cell, looks, steady, ms]) => {
      const w = window as any
      const textOf = (c: string) => document.querySelector(`[data-cell="${c}"]`)?.textContent ?? ''
      const bg = (c: string) => {
        const el = document.querySelector(`[data-cell="${c}"]`)
        return el ? getComputedStyle(el).backgroundColor : ''
      }
      const before = textOf(cell)
      const colours = looks.map(bg)
      w.__watch = { enter: 0, shown: 0, look: 0, blank: 0, done: false }
      addEventListener('keydown', (e) => e.key === 'Enter' && (w.__watch.enter = performance.now()), { capture: true, once: true })
      const check = () => {
        const x = w.__watch
        const now = performance.now()
        if (x.enter) {
          // a cell drawn empty (its window dropped) has neither its new text nor its new colour yet
          const text = textOf(cell)
          if (!x.shown && text.trim() && text !== before) x.shown = now
          if (!x.look && looks.some((c, i) => textOf(c).trim() && bg(c) !== colours[i])) x.look = now
          if (!textOf(steady).trim()) x.blank++
          if (now - x.enter > ms && x.shown) {
            x.done = true
            return
          }
        }
        requestAnimationFrame(check)
      }
      requestAnimationFrame(check)
    },
    [watch, looks, steady, WATCH_MS] as const,
  )
  await page.keyboard.press('Enter')
  await page.waitForFunction(() => (window as any).__watch.done, null, { timeout: 30_000 })
  return page.evaluate(() => {
    const x = (window as any).__watch
    return { shown: x.shown - x.enter, look: x.look ? x.look - x.enter : null, blank: x.blank }
  })
}

async function lastEdit(page: Page, count: number): Promise<{ ms: number; changes: number }> {
  await page.waitForFunction((n) => (window as any).__xs.edits.length >= n, count, { timeout: 30_000 })
  return page.evaluate((n) => (window as any).__xs.edits[n - 1], count)
}

/** Wait until a cell (`row:col`) shows `text`. */
async function shows(page: Page, cell: string, text: string): Promise<void> {
  await page.waitForFunction(([c, t]) => document.querySelector(`[data-cell="${c}"]`)?.textContent?.trim() === t, [cell, text] as const, { timeout: 30_000 })
}

/** The row numbers drawn, in order. */
const numbers = (page: Page) => page.evaluate(() => [...document.querySelectorAll('.xs-num')].map((n) => n.textContent!.trim()))

/** Excel's outline in the pool: a collapsed section shown by its button and hidden again by Control+Z, a level button, the columns' group. */
async function outlineScenario(page: Page, base: string): Promise<void> {
  await open(page, base, 'Структура!A1', '1:1')
  const rows = async (want: string) => {
    await page.waitForFunction((w) => [...document.querySelectorAll('.xs-num')].map((n) => n.textContent!.trim()).slice(0, 8).join(' ') === w, want, { timeout: 10_000 })
  }
  await rows('1 2 3 4 5 9 10 11')
  await page.locator('.xs-ol-rows [data-xs-outline="6:8:2"]').click()
  await rows('1 2 3 4 5 6 7 8')
  await page.screenshot({ path: resolve(out, 'local-1920-outline.png') })
  await page.locator('.xs-scroller').focus()
  await page.keyboard.press('Control+Z')
  await rows('1 2 3 4 5 9 10 11')
  await page.locator('.xs-ol-levels-rows [data-xs-outline-level="1"]').click()
  await rows('1 10 11 12 13 14 15 16')
  await page.locator('.xs-ol-columns [data-xs-outline="2:4:1"]').click()
  await page.waitForFunction(() => [...document.querySelectorAll('.xs-head')].map((h) => h.textContent!.trim()).slice(0, 2).join(' ') === 'A E')
  const total = await page.locator('[data-cell="10:5"]').textContent()
  await page.screenshot({ path: resolve(out, 'local-1920-outline-level.png') })
  console.log(`[1920] outline: + shows 6:8, Control+Z hides it, level 1 leaves the total (E10 = ${total?.trim()}), columns B:D collapsed; rows now ${(await numbers(page)).slice(0, 4).join(' ')}`)
}

/**
 * Data validation through the pool: a value a stop rule and one a warning rule do not take are refused with the rule's text
 * on the message line (the sheet has no alert yet), and the cells keep their values.
 */
async function validationScenario(page: Page, base: string): Promise<void> {
  for (const [cell, at, typed, text, severity, button, after] of [
    ['B4', '4:2', '500', 'Целое число от 1 до 100', 'stop', 'cancel', '5'],
    ['B5', '5:2', '50', 'Скидка больше 30% требует согласования', 'warning', 'yes', '50%'],
  ] as const) {
    await open(page, base, `Проверка!${cell}`, at)
    await page.locator('.xs-scroller').focus()
    await page.keyboard.type(typed)
    await page.keyboard.press('Enter')
    await page.waitForFunction(([s, t]) => document.querySelector(`[data-xs-alert="${s}"]`)?.textContent?.includes(t), [severity, text] as const, { timeout: 10_000 })
    if (severity === 'stop') await page.screenshot({ path: resolve(out, 'local-1920-validation.png') })
    await page.locator(`[data-xs-alert-button="${button}"]`).click()
    await shows(page, at, after)
  }
  console.log('[1920] validation: 500 under a stop rule refused through the pool, Cancel kept 5; 50% under a warning written on Yes')
}

/** A small ru-RU file opened from the field, typed into, filled with Control+Return, pasted into, undone, redone, saved and opened again. */
async function fileScenario(browser: Browser, base: string): Promise<void> {
  const file = resolve(out, 'local-small.xlsx')
  registerLocale(ruRU)
  const book = { locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 2 }, A2: { value: 3 }, A3: { value: 4 }, C1: { numberFormat: '0%' } } }] }
  writeFileSync(file, (await writeXlsx({ book })).data)
  const context = await browser.newContext({ viewport: { width: 1920, height: 1200 }, acceptDownloads: true })
  const page = await context.newPage()
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  const shot = (name: string) => page.screenshot({ path: resolve(out, `local-file-${name}.png`) })
  await page.goto(`${base}local.html?rows=34`)
  await page.locator('[data-open]').setInputFiles(file)
  await shows(page, '1:1', '2')
  // Typing over a cell.
  await page.locator('[data-cell="1:4"]').click()
  await page.keyboard.type('abc')
  await page.keyboard.press('Enter')
  await shows(page, '1:4', 'abc')
  // Control+Return over B1:B3: the formula moves as a copy moves it.
  await page.locator('[data-cell="1:2"]').click()
  await page.locator('[data-cell="3:2"]').click({ modifiers: ['Shift'] })
  await page.keyboard.type('=A1*2')
  await page.keyboard.press('Control+Enter')
  await shows(page, '3:2', '8')
  await page.locator('[data-cell="3:2"]').click()
  const bar = (await page.locator('.xs-fx').textContent())?.trim()
  if (bar !== '=A3*2') throw new Error(`B3 after Control+Return: ${bar}`)
  await shot('fill')
  // A paste at C1, a percent cell: 5 is 500 % (no percent editor), '007 keeps its apostrophe.
  await page.locator('[data-cell="1:3"]').click()
  await page.locator('.xs-scroller').evaluate((el) => {
    const data = new DataTransfer()
    data.setData('text/plain', "5\t'007")
    el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }))
  })
  await shows(page, '1:3', '500%')
  await shows(page, '1:4', "'007")
  await shot('paste')
  // One undo takes the paste back whole, a redo does it again.
  await page.keyboard.press('Control+Z')
  await shows(page, '1:3', '')
  await shows(page, '1:4', 'abc')
  await page.keyboard.press('Control+Y')
  await shows(page, '1:4', "'007")
  // Saved and opened again: the values come back.
  const [download] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: 'Сохранить xlsx' }).click()])
  const saved = resolve(out, 'local-small-saved.xlsx')
  await download.saveAs(saved)
  await page.locator('[data-open]').setInputFiles(saved)
  await shows(page, '3:2', '8')
  await shows(page, '1:3', '500%')
  await shows(page, '1:4', "'007")
  await shot('reopened')
  if (errors.length) throw new Error(`page errors: ${errors.join('; ')}`)
  console.log('[file] typed, Control+Return, paste, undo, redo, saved and opened again: values as edited')
  await context.close()
}

async function production(browser: Browser): Promise<void> {
  await build({ configFile: config, logLevel: 'error' })
  const server = await preview({ configFile: config, preview: { port: 0 }, logLevel: 'error' })
  const base = server.resolvedUrls!.local[0]!
  try {
    for (const size of [
      { name: '1920', width: 1920, height: 1200, phone: false },
      { name: '390', width: 390, height: 844, phone: true },
    ]) {
      const context = await browser.newContext({ viewport: { width: size.width, height: size.height }, deviceScaleFactor: size.phone ? 2 : 1, isMobile: size.phone, hasTouch: size.phone })
      const page = await context.newPage()
      const errors: string[] = []
      page.on('pageerror', (e) => errors.push(e.message))
      page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))

      const opened: Opened[] = []
      for (let i = 0; i < (size.phone ? 1 : RUNS); i++) opened.push(await open(page, base, 'Витрина!B3', '3:2'))
      const o = opened[0]!
      if (o.workers < 1) throw new Error('no Web Worker computes the book')
      console.log(`[${size.name}] open ${ROWS} rows: first cells drawn ${fmt(opened.map((x) => x.drawn))}; session ready ${fmt(opened.map((x) => x.ready))}; book built ${fmt(opened.map((x) => x.built))}; workers ${o.workers}`)
      await page.waitForTimeout(300)
      await page.screenshot({ path: resolve(out, `local-${size.name}-open.png`) })

      // A value in the table: its total below on screen, and the colour scale over the shares recoloured (the edit's repaint).
      const share = () => page.evaluate(() => getComputedStyle(document.querySelector('[data-cell="4:4"]')!).backgroundColor)
      const before = await share()
      const watched: Watched[] = []
      for (const v of size.phone ? ['2500'] : ['2500', '3750,5', '1 000', '42', '7,25']) {
        watched.push(await typeAndWatch(page, v, '9:2', size.phone, ['3:4', '4:4', '5:4', '6:4', '7:4']))
        if (watched.length === 1) await page.waitForFunction((c) => getComputedStyle(document.querySelector('[data-cell="4:4"]')!).backgroundColor !== c, before, { timeout: 10_000 })
        // Enter moved down a row: back to B3
        if (!size.phone) await page.keyboard.press('ArrowUp')
      }
      console.log(`[${size.name}] the share D4 recoloured by the edit: ${before} → ${await share()}`)
      const total = await page.locator('[data-cell="9:2"]').textContent()
      console.log(`[${size.name}] edit Витрина!B3, Enter to the total B9 on screen: ${fmt(watched.map((x) => x.shown))}; B9 = ${total}`)
      const looks = watched.flatMap((x) => (x.look === null ? [] : [x.look]))
      console.log(`[${size.name}] Enter to the shares D3:D7 recoloured: ${looks.length ? fmt(looks) : 'none'} (${looks.length} of ${watched.length} edits recoloured)`)
      const blank = watched.map((x) => x.blank)
      console.log(`[${size.name}] frames drawn with A3 empty within ${WATCH_MS} ms of Enter: ${blank.join(', ')}`)
      if (blank.some((n) => n > 0)) throw new Error(`an edit drew the sheet empty: ${blank.join(', ')} frames`)
      await page.waitForTimeout(200)
      await page.screenshot({ path: resolve(out, `local-${size.name}-edit.png`) })

      if (!size.phone) {
        // An amount on the data sheet: the SUMIFS of other sheets recompute.
        await open(page, base, 'Данные!D2', '2:4')
        const before = await page.evaluate(() => (window as any).__xs.edits.length as number)
        await page.locator('.xs-scroller').focus()
        await page.keyboard.type('123456,78')
        await page.keyboard.press('Enter')
        const e = await lastEdit(page, before + 1)
        console.log(`[${size.name}] edit Данные!D2: ${Math.round(e.ms)} ms in the source, ${e.changes} cells changed`)
        await page.waitForTimeout(200)
        await page.screenshot({ path: resolve(out, `local-${size.name}-data.png`) })
        // A formula that does not parse: the field opens again with the message.
        await open(page, base, 'Витрина!F3', '3:1')
        await page.locator('.xs-scroller').focus()
        await page.keyboard.type('=СУММ(B3:B4')
        await page.keyboard.press('Enter')
        await page.waitForTimeout(300)
        await page.screenshot({ path: resolve(out, `local-${size.name}-invalid.png`) })
        await page.keyboard.press('Escape')
        await outlineScenario(page, base)
        await validationScenario(page, base)
        const bytes = await page.evaluate(() => (window as any).__xs.save() as Promise<number>)
        if (!(bytes > 1000)) throw new Error(`saved xlsx of ${bytes} bytes`)
        console.log(`[${size.name}] saved xlsx: ${bytes} bytes`)
      }
      if (errors.length) throw new Error(`page errors: ${errors.join('; ')}`)
      await context.close()
    }
    await fileScenario(browser, base)
  } finally {
    await new Promise<void>((done) => server.httpServer.close(() => done()))
  }
}

/** The dev server serves xlcalc pre-bundled: the pool must still find its thread there. */
async function dev(browser: Browser): Promise<void> {
  const server = await createServer({ configFile: config, server: { port: 0 }, logLevel: 'error' })
  await server.listen()
  try {
    const page = await browser.newPage({ viewport: { width: 1920, height: 1200 } })
    const o = await open(page, server.resolvedUrls!.local[0]!, 'Витрина!B3', '3:2')
    if (o.workers < 1) throw new Error('dev: no Web Worker computes the book')
    console.log(`[dev] first cells drawn ${Math.round(o.drawn)} ms, workers ${o.workers}`)
    await page.close()
  } finally {
    await server.close()
  }
}

const browser = await chromium.launch()
try {
  await production(browser)
  await dev(browser)
} finally {
  await browser.close()
}
console.log(`screens in ${out}`)
