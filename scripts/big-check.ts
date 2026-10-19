// Large workbooks in the Excel-like app: the example's spa.html built for production and served by `vite preview`, driven
// by Playwright in Chromium at 1920 × 1200 with the interface in Russian. Each workbook is opened from the file input and
// measured: the first drawn cells of the sheet the app opens and Ready, a jump to the largest sheet, scrolling it down
// and right screen by screen, a number typed into a cell that another on-screen cell depends on, until the dependent displays
// its new value, Ctrl+Z, Save as (time and size; the saved file read back has the workbook's sheets and formulas), the heap
// of the page and of the workbook's Web Worker, the longest pause between the page's frames, the workbook's language changed
// to en-US in its settings (it reopens on the same sheet), and page errors.
// BIG_URL=<url> measures a published app instead of the built one (synthetic workbooks only: what it opens goes to the published app).
// Workbooks: BIG_XLSX, paths separated by ':', read in place (nothing is copied, the saved file is deleted); without it,
// xlcalc's synthetic workbook of 34,000 rows, written to test-results. BIG_PROFILE=1 also writes the page's CPU profiles
// of the opening, the jump, the scrolling and the saving to test-results. Run: npm run check:big. Output: a table on stdout, test-results/big-check.json and
// test-results/screens/big-*.png.
import { mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, resolve } from 'node:path'
import { chromium, type Browser, type CDPSession, type Page } from 'playwright'
import { build, preview } from 'vite'
import { compute, type BookInput, type CellInput, type SheetInput } from 'xlcalc'
import { cellAddress, parseCell, parseRange } from 'xlcalc/a1'
import { quoteSheetIfNeeded, scanReferences } from 'xlcalc/text'
import { synthBook } from 'xlcalc/unstable/synth'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const root = resolve(import.meta.dirname, '..')
const config = resolve(root, 'vite.example.config.ts')
const results = resolve(root, 'test-results')
const screens = resolve(results, 'screens')
mkdirSync(screens, { recursive: true })
const SYNTH_ROWS = 34_000
const PROFILE = process.env.BIG_PROFILE === '1'
/** Screens scrolled down and right on the largest sheet. */
const DOWN = 5
const RIGHT = 3
const LONG = 600_000

// ── the workbook, examined in Node: the sheet the app should open, the largest sheet and cells to edit ──

/** A number to type into `cell` of `sheet`, and a nearby cell whose formula refers to it. */
interface Edit {
  sheet: string
  cell: string
  value: number
  dependent: string
}
interface Plan {
  label: string
  path: string
  bytes: number
  cells: number
  formulas: number
  sheets: string[]
  /** The locale the file specifies (xlsx stores none unless xlcalc wrote it). */
  locale: string | undefined
  /** The sheet the app opens, as Excel opens the file: the saved active one, otherwise the first visible one. */
  first: string | undefined
  largest: { name: string; content: number }
  edits: Edit[]
}

const hasContent = (c: CellInput) => c.formula !== undefined || (c.value !== undefined && c.value !== null && c.value !== '')

/** Hidden rows or columns of a sheet, by number, from its spans (`5:9`, `B:D`); a size of 0 also hides a row or column. */
function hiddenLines(sheet: SheetInput, axis: 'rows' | 'columns'): Set<number> {
  const out = new Set<number>()
  const lines = (axis === 'rows' ? sheet.rows : sheet.columns) ?? {}
  for (const [key, line] of Object.entries(lines)) {
    const size = 'height' in line ? line.height : 'width' in line ? line.width : undefined
    if (!line.hidden && size !== 0) continue
    const r = parseRange(key.includes(':') ? key : `${key}:${key}`)
    if (!r) continue
    const [a, b] = axis === 'rows' ? [r.top, r.bottom] : [r.left, r.right]
    for (let i = a; i <= b; i++) out.add(i)
  }
  return out
}

/**
 * Cells to edit on a sheet: a manually entered number (no formula) that a formula within 20 rows and 8 columns refers to
 * as a single cell on the same sheet, both on visible rows and columns and outside merged ranges; the shortest formulas
 * first, at most `count`, each with its own input cell. `numeric` says whether a formula's cell contains a number now.
 */
function editsOn(sheet: SheetInput, numeric: (addr: string) => boolean, count: number): Edit[] {
  const rowsHidden = hiddenLines(sheet, 'rows')
  const colsHidden = hiddenLines(sheet, 'columns')
  const merged = (sheet.merges ?? []).flatMap((m) => {
    const r = parseRange(m)
    return r ? [r] : []
  })
  const shown = (row: number, col: number) => !rowsHidden.has(row) && !colsHidden.has(col) && !merged.some((m) => row >= m.top && row <= m.bottom && col >= m.left && col <= m.right)
  const found: (Edit & { length: number })[] = []
  for (const [addr, cell] of Object.entries(sheet.cells)) {
    const f = cell.formula
    if (!f || f.length > 80 || f.includes('[') || !numeric(addr)) continue
    const at = parseCell(addr)
    if (!at || !shown(at.row, at.col)) continue
    for (const r of scanReferences(f)) {
      if (r.kind !== 'cells' || r.sheet !== undefined || r.spill || r.top !== r.bottom || r.left !== r.right) continue
      if (Math.abs(r.top - at.row) > 20 || Math.abs(r.left - at.col) > 8 || !shown(r.top, r.left)) continue
      const input = cellAddress(r.top, r.left)
      const value = sheet.cells[input]?.value
      if (sheet.cells[input]?.formula !== undefined || typeof value !== 'number') continue
      found.push({ sheet: sheet.name, cell: input, value, dependent: addr, length: f.length })
      break
    }
  }
  found.sort((a, b) => a.length - b.length)
  const seen = new Set<string>()
  const out: Edit[] = []
  for (const { length: _, ...e } of found) {
    if (seen.has(e.cell)) continue
    seen.add(e.cell)
    out.push(e)
    if (out.length === count) break
  }
  return out
}

/** What the check needs of a workbook; `values` returns a formula's value as the file stored it or the engine calculated it. */
function planOf(label: string, path: string, book: BookInput, values: (sheet: number, addr: string) => unknown): Plan {
  let cells = 0
  let formulas = 0
  const content = book.sheets.map((s) => {
    const all = Object.values(s.cells)
    cells += all.length
    formulas += all.filter((c) => c.formula !== undefined).length
    return all.filter(hasContent).length
  })
  const visible = book.sheets.map((s, i) => ({ s, i })).filter(({ s }) => !s.hidden)
  const first = book.activeSheet ?? visible[0]?.s.name
  const big = visible.reduce((a, b) => (content[b.i]! > content[a.i]! ? b : a), visible[0]!)
  // the largest sheet first, then the others by size, until three cells are found
  const order = [...visible].sort((a, b) => (a.i === big.i ? -1 : b.i === big.i ? 1 : content[b.i]! - content[a.i]!))
  const edits: Edit[] = []
  for (const { s, i } of order) {
    if (edits.length >= 3) break
    edits.push(...editsOn(s, (addr) => typeof values(i, addr) === 'number', 3 - edits.length))
  }
  return { label, path, bytes: statSync(path).size, cells, formulas, sheets: book.sheets.map((s) => s.name), locale: book.locale, first, largest: { name: big.s.name, content: content[big.i]! }, edits }
}

async function planFile(path: string): Promise<Plan> {
  const read = await readXlsx(new Uint8Array(readFileSync(path)))
  return planOf(basename(path), path, read.book, (i, addr) => read.cached.sheets[i]?.cells[addr]?.value)
}

/** The synthetic workbook written with its values; a cell beyond its data sheet's last column refers to that column's first amount, so an edit there shows on screen. */
async function synthPlan(): Promise<Plan> {
  const { book } = synthBook({ seed: 7, rows: SYNTH_ROWS })
  const data = book.sheets.find((s) => s.name === 'Данные')!
  const cols = Object.keys(data.cells).reduce((m, a) => Math.max(m, parseCell(a)?.col ?? 0), 0)
  data.cells[cellAddress(2, cols + 2)] = { formula: `${cellAddress(2, cols)}*2` }
  const result = compute(book)
  const path = resolve(results, `big-synth-${SYNTH_ROWS}.xlsx`)
  writeFileSync(path, (await writeXlsx({ book, result })).data)
  return planOf(`synth ${SYNTH_ROWS}`, path, book, (i, addr) => result.sheets[i]?.cells[addr]?.value)
}

// ── the page: frames, long tasks and waits measured where they happen ──

/** Installed before the app loads: long tasks, frame gaps over 50 ms, and a wait for a condition two frames in a row. */
function probe(): void {
  const w = window as unknown as { __big: Record<string, unknown> }
  const long: [number, number][] = []
  const gaps: [number, number][] = []
  try {
    new PerformanceObserver((list) => {
      for (const e of list.getEntries()) long.push([e.startTime, e.duration])
    }).observe({ type: 'longtask', buffered: true })
  } catch {
    // no long tasks in this browser: frame gaps still reveal the pauses
  }
  let last = performance.now()
  const frame = (t: number) => {
    if (t - last > 50) gaps.push([last, t - last])
    last = t
    requestAnimationFrame(frame)
  }
  requestAnimationFrame(frame)
  const q = (s: string) => document.querySelector(s)
  const tab = () => q('.xs-tab-on')?.textContent?.trim() ?? ''
  const drawn = (sheet?: string) => (!sheet || tab() === sheet) && !!q('.xs-table, .xs-empty') && !q('.xs-loading')
  /** The time of the first of two consecutive frames where `test` holds, after at least `min` frames. */
  const until = (test: () => boolean, min = 1, timeout = 600_000) =>
    new Promise<number>((done, fail) => {
      const t0 = performance.now()
      let n = 0
      let first = 0
      const f = () => {
        n++
        if (n >= min && test()) {
          if (first) return done(first)
          first = performance.now()
        } else first = 0
        if (performance.now() - t0 > timeout) return fail(new Error(`no change in ${timeout} ms`))
        requestAnimationFrame(f)
      }
      requestAnimationFrame(f)
    })
  /** The longest frame gap and long task that overlap [from, to]. */
  const pause = (from: number, to: number) => {
    const over = (xs: [number, number][]) => Math.max(0, ...xs.filter(([s, d]) => s < to && s + d > from).map(([, d]) => d))
    return { frame: over(gaps), task: over(long) }
  }
  w.__big = { tab, drawn, until, pause }
}

interface Big {
  tab: () => string
  drawn: (sheet?: string) => boolean
  until: (test: () => boolean, min?: number, timeout?: number) => Promise<number>
  pause: (from: number, to: number) => { frame: number; task: number }
}
declare const __big: Big
const now = (page: Page) => page.evaluate(() => performance.now())
const pauseOf = (page: Page, from: number, to: number) => page.evaluate(([a, b]) => __big.pause(a, b), [from, to] as const)
const idle = (page: Page) => page.evaluate(() => __big.until(() => !document.querySelector('.xs-loading'), 2))
const ready = (page: Page) => page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: LONG })

/** A non-flattened CDP session to a target (a Web Worker), over the browser's session; a call that an ended thread does not answer returns undefined after 10 s. */
async function attach(cdp: CDPSession, targetId: string) {
  const { sessionId } = await cdp.send('Target.attachToTarget', { targetId, flatten: false })
  let id = 0
  const call = (method: string, params: object = {}) =>
    new Promise<any>((done) => {
      const mine = ++id
      const on = (e: { sessionId: string; message: string }) => {
        if (e.sessionId !== sessionId) return
        const msg = JSON.parse(e.message) as { id?: number; result?: unknown; error?: unknown }
        if (msg.id !== mine) return
        finish(msg.result ?? msg.error)
      }
      const timer = setTimeout(() => finish(undefined), 10_000)
      const finish = (v: unknown) => {
        clearTimeout(timer)
        cdp.off('Target.receivedMessageFromTarget', on)
        done(v)
      }
      cdp.on('Target.receivedMessageFromTarget', on)
      cdp.send('Target.sendMessageToTarget', { sessionId, message: JSON.stringify({ id: mine, method, params }) }).catch(() => finish(undefined))
    })
  return { call, detach: () => cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {}) }
}

/** The page's CPU profile of a step, written as test-results/big-<n>-<name>.cpuprofile when BIG_PROFILE=1. */
async function profiled<T>(page: Page, n: number, name: string, step: () => Promise<T>): Promise<T> {
  if (!PROFILE) return step()
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Profiler.enable')
  await cdp.send('Profiler.start')
  try {
    return await step()
  } finally {
    const { profile } = await cdp.send('Profiler.stop')
    writeFileSync(resolve(results, `big-${n}-${name}.cpuprofile`), JSON.stringify(profile))
    await cdp.detach()
  }
}

/** Heap in use after a garbage collection, in MB: the page's, and its Web Workers' together. */
async function heaps(page: Page, browserCdp: CDPSession): Promise<{ page: number; worker: number }> {
  const ps = await page.context().newCDPSession(page)
  await ps.send('HeapProfiler.collectGarbage')
  const { usedSize } = await ps.send('Runtime.getHeapUsage')
  const { targetInfo } = await ps.send('Target.getTargetInfo')
  await ps.detach()
  const { targetInfos } = await browserCdp.send('Target.getTargets')
  let worker = 0
  for (const t of targetInfos.filter((x) => x.type === 'worker' && x.browserContextId === targetInfo.browserContextId)) {
    // a thread that ended between the listing and the attach has no heap to count
    const w = await attach(browserCdp, t.targetId).catch(() => null)
    if (!w) continue
    await w.call('HeapProfiler.collectGarbage')
    worker += ((await w.call('Runtime.getHeapUsage')) as { usedSize?: number } | undefined)?.usedSize ?? 0
    await w.detach()
  }
  return { page: usedSize / 2 ** 20, worker: worker / 2 ** 20 }
}

/** A location through the Name Box, as a user types it; the grid has focus afterwards. */
async function go(page: Page, place: string): Promise<void> {
  await page.locator('.xs-address').fill(place)
  await page.keyboard.press('Enter')
  await idle(page)
}

const cellKey = (addr: string) => {
  const at = parseCell(addr)!
  return `${at.row}:${at.col}`
}
/** A cell's drawn text, or null when it is not drawn inside the scroller's viewport. */
const shownText = (page: Page, key: string) =>
  page.evaluate((k) => {
    const el = document.querySelector<HTMLElement>(`[data-cell="${k}"]`)
    const frame = document.querySelector<HTMLElement>('.xs-scroller')?.getBoundingClientRect()
    if (!el || !frame) return null
    const r = el.getBoundingClientRect()
    return r.top >= frame.top && r.bottom <= frame.bottom && r.left >= frame.left && r.right <= frame.right ? (el.textContent ?? '').trim() : null
  }, key)

/** Presses a key and measures, from its keydown, until `key`'s cell displays text other than `before` (or `back` again). */
async function pressAndWatch(page: Page, press: string, key: string, before: string, back: boolean): Promise<number> {
  await page.evaluate(
    ([k, b, toBack]) => {
      const w = window as unknown as { __watch: { down: number; shown: number } }
      w.__watch = { down: 0, shown: 0 }
      addEventListener('keydown', () => (w.__watch.down = performance.now()), { capture: true, once: true })
      const check = () => {
        const now = document.querySelector(`[data-cell="${k}"]`)?.textContent?.trim()
        if (w.__watch.down && (toBack ? now === b : now !== undefined && now !== b)) w.__watch.shown = performance.now()
        else requestAnimationFrame(check)
      }
      requestAnimationFrame(check)
    },
    [key, before, back] as const,
  )
  await page.keyboard.press(press)
  await page.waitForFunction(() => (window as unknown as { __watch: { shown: number } }).__watch.shown > 0, null, { timeout: 120_000 })
  return page.evaluate(() => {
    const w = (window as unknown as { __watch: { down: number; shown: number } }).__watch
    return w.shown - w.down
  })
}

interface Row {
  label: string
  mb: number
  cells: number
  formulas: number
  locale: string
  sheet: string
  drawn: number
  ready: number
  read?: number
  openPause: { frame: number; task: number }
  largest: string
  jump: number
  down: number[]
  right: number[]
  scrollPause: { frame: number; task: number }
  topRow: number
  edit?: { cell: string; dependent: string; shown: number; ready: number; undo: number }
  editTried: number
  save: number
  saveMb: number
  savedFormulas: number
  savePause: { frame: number; task: number }
  heapOpen: { page: number; worker: number }
  heapEnd: { page: number; worker: number }
  relocale: number
  errors: string[]
  failures: string[]
}

async function check(browser: Browser, base: string, plan: Plan, n: number): Promise<Row> {
  const context = await browser.newContext({ viewport: { width: 1920, height: 1200 }, acceptDownloads: true })
  await context.addInitScript(probe)
  const page = await context.newPage()
  const browserCdp = await browser.newBrowserCDPSession()
  const errors: string[] = []
  const failures: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))
  const shot = (name: string) => page.screenshot({ path: resolve(screens, `big-${n}-${name}.png`) })
  const name = basename(plan.path).replace(/\.xlsx$/i, '')

  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await ready(page)

  // ── open: from the file input's change to the first drawn cells of the sheet the app should open, and to Ready ──
  await page.evaluate(
    ([want, sheet]) => {
      const w = window as unknown as { __open: { change: number; any: number; drawn: number; ready: number } }
      const o = (w.__open = { change: 0, any: 0, drawn: 0, ready: 0 })
      document.querySelector('[data-spa-file]')!.addEventListener('change', () => (o.change = performance.now()), { capture: true, once: true })
      const named = () => (document.querySelector<HTMLInputElement>('[data-spa-name]')?.value ?? '') === want
      const f = () => {
        const t = performance.now()
        if (o.change && named()) {
          if (!o.any && __big.drawn()) o.any = t
          if (!o.drawn && __big.drawn(sheet)) o.drawn = t
          if (!o.ready && document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово') o.ready = t
        }
        if (o.ready && (o.drawn || t - o.ready > 3000)) return
        requestAnimationFrame(f)
      }
      requestAnimationFrame(f)
    },
    [name, plan.first ?? ''] as const,
  )
  await profiled(page, n, 'open', async () => {
    await page.locator('[data-spa-file]').setInputFiles(plan.path)
    await page.waitForFunction(
      () => {
        const o = (window as unknown as { __open: { ready: number; drawn: number } }).__open
        return o.ready > 0 && (o.drawn > 0 || performance.now() - o.ready > 3000)
      },
      null,
      { timeout: LONG, polling: 500 },
    )
  })
  const o = await page.evaluate(() => (window as unknown as { __open: { change: number; any: number; drawn: number; ready: number } }).__open)
  const opened = await page.evaluate(() => __big.tab())
  if (plan.first && opened !== plan.first) failures.push(`opened ${opened}, Excel opens ${plan.first}`)
  const locale = ((await page.locator('.spa-meta').textContent()) ?? '').split('·')[0]!.trim()
  if (!plan.locale && locale !== 'ru-RU') failures.push(`a file without a language opened in ${locale}, the interface is Russian`)
  const timings = await page.evaluate(() => (window as unknown as { __spa: { timings: () => Record<string, number> } }).__spa.timings())
  const openPause = await pauseOf(page, o.change, o.ready)
  await page.waitForTimeout(300)
  await shot('open')
  const heapOpen = await heaps(page, browserCdp)

  // ── the largest sheet, from a click on its tab ──
  let jump = 0
  if (plan.largest.name !== (await page.evaluate(() => __big.tab()))) {
    const click = (sheet: string) =>
      page.evaluate(async (name) => {
        const tab = [...document.querySelectorAll<HTMLElement>('.xs-tab')].find((t) => t.textContent?.trim() === name)!
        tab.scrollIntoView()
        const t0 = performance.now()
        tab.click()
        return (await __big.until(() => __big.drawn(name))) - t0
      }, sheet)
    jump = await profiled(page, n, 'jump', () => click(plan.largest.name))
  }
  await shot('largest')

  // ── scrolling, a screen at a time: until the windows have loaded and are drawn ──
  const scrollFrom = await now(page)
  const scroll = (axis: 'down' | 'right', count: number) =>
    page.evaluate(
      async ([a, c]) => {
        const el = document.querySelector<HTMLElement>('.xs-scroller')!
        const out: number[] = []
        for (let i = 0; i < c; i++) {
          const t0 = performance.now()
          if (a === 'down') el.scrollTop += el.clientHeight - 40
          else el.scrollLeft += el.clientWidth - 80
          out.push((await __big.until(() => !document.querySelector('.xs-loading'), 2)) - t0)
        }
        return out
      },
      [axis, count] as const,
    )
  const down = await profiled(page, n, 'down', () => scroll('down', DOWN))
  await shot('down')
  const right = await profiled(page, n, 'right', () => scroll('right', RIGHT))
  await shot('right')
  const topRow = await page.evaluate(() => Number(document.querySelector<HTMLElement>('tbody th.xs-num:not(.xs-fz-corner)')?.dataset.row ?? 0))
  const scrollPause = await pauseOf(page, scrollFrom, await now(page))

  // ── an edit with its dependent on screen, then Ctrl+Z ──
  let edit: Row['edit']
  let tried = 0
  for (const e of plan.edits) {
    tried++
    const place = (addr: string) => `${quoteSheetIfNeeded(e.sheet)}!${addr}`
    await go(page, place(e.dependent))
    await go(page, place(e.cell))
    const dep = cellKey(e.dependent)
    const before = await shownText(page, dep)
    const own = await shownText(page, cellKey(e.cell))
    if (before === null || own === null) continue
    let typed = Math.round(e.value * 3) || 7
    if (typed === e.value) typed += 7
    await page.keyboard.type(String(typed))
    let shown: number
    try {
      shown = await pressAndWatch(page, 'Enter', dep, before, false)
    } catch {
      // the dependent did not change: undo the edit and try the next cell
      await ready(page)
      await page.keyboard.press('Control+Z')
      await ready(page)
      continue
    }
    const t0 = await now(page)
    await ready(page)
    const readyMs = (await now(page)) - t0 + shown
    await page.waitForTimeout(200)
    await shot('edit')
    const undo = await pressAndWatch(page, 'Control+Z', dep, before, true)
    await ready(page)
    edit = { cell: `${e.sheet}!${e.cell}`, dependent: e.dependent, shown, ready: readyMs, undo }
    break
  }
  if (!edit) failures.push(`no edit showed on screen (${plan.edits.length} cells found, ${tried} tried)`)

  // ── Save as: the app's own time from the click to the download, and the file's size; the file is deleted ──
  const saveFrom = await now(page)
  // the download, or the app's message explaining why there is none
  const failed = page.waitForSelector('[data-spa-note]', { timeout: LONG }).then(async (n) => {
    throw new Error(`Save as: ${await n.textContent()}`)
  })
  failed.catch(() => {})
  const [download] = await profiled(page, n, 'save', () => Promise.all([Promise.race([page.waitForEvent('download', { timeout: LONG }), failed]), page.locator('[data-spa="save"]').click()]))
  const saveTo = await now(page)
  const file = await download.path()
  const saveMb = statSync(file).size / 2 ** 20
  // the saved file, read back where Playwright put it, then deleted
  const back = (await readXlsx(new Uint8Array(readFileSync(file)))).book
  await download.delete()
  const savedFormulas = back.sheets.reduce((n, s) => n + Object.values(s.cells).filter((c) => c.formula !== undefined).length, 0)
  if (back.sheets.map((s) => s.name).join('\n') !== plan.sheets.join('\n')) failures.push('the saved file has other sheets than the book')
  if (savedFormulas !== plan.formulas) failures.push(`the saved file has ${savedFormulas} formulas, the book ${plan.formulas}`)
  await ready(page)
  const after = await page.evaluate(() => (window as unknown as { __spa: { timings: () => Record<string, number> } }).__spa.timings())
  const savePause = await pauseOf(page, saveFrom, saveTo)
  const heapEnd = await heaps(page, browserCdp)

  // ── the workbook's language from its settings: the workbook reopens in en-US, on the same sheet ──
  const sheetBefore = await page.evaluate(() => __big.tab())
  await page.locator('[data-spa="settings"]').click()
  await page.locator('[data-spa-book-locale]').selectOption('en-US')
  const relocaleFrom = await now(page)
  await page.locator('[data-spa="apply"]').click()
  await page.waitForFunction(() => document.querySelector('.spa-meta')?.textContent?.trim().startsWith('en-US'), null, { timeout: LONG })
  await ready(page)
  const relocale = (await page.evaluate((sheet) => __big.until(() => __big.drawn(sheet)), sheetBefore)) - relocaleFrom
  const relocaled = await page.evaluate(() => __big.tab())
  if (relocaled !== sheetBefore) failures.push(`the book in en-US opened on ${relocaled}, it stood on ${sheetBefore}`)
  await shot('en-US')
  await browserCdp.detach()
  await context.close()
  return {
    label: plan.label,
    mb: plan.bytes / 2 ** 20,
    cells: plan.cells,
    formulas: plan.formulas,
    locale,
    sheet: opened,
    drawn: (o.drawn || o.any) - o.change,
    ready: o.ready - o.change,
    ...(timings.read ? { read: timings.read } : {}),
    openPause,
    largest: plan.largest.name,
    jump,
    down,
    right,
    scrollPause,
    topRow,
    ...(edit ? { edit } : {}),
    editTried: tried,
    save: after.save || saveTo - saveFrom,
    saveMb,
    savedFormulas,
    savePause,
    heapOpen,
    heapEnd,
    relocale,
    errors,
    failures,
  }
}

// ── the table ──
const ms = (x: number) => `${Math.round(x)}`
const sec = (x: number) => (x >= 1000 ? `${(x / 1000).toFixed(1)} s` : `${Math.round(x)} ms`)
const med = (xs: number[]) => (xs.length ? [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)]! : 0)
function print(rows: Row[]): void {
  const lines = [
    ['', ...rows.map((r) => r.label)],
    ['file, MB', ...rows.map((r) => r.mb.toFixed(1))],
    ['cells / formulas', ...rows.map((r) => `${r.cells} / ${r.formulas}`)],
    ['book locale', ...rows.map((r) => r.locale)],
    ['sheet opened', ...rows.map((r) => r.sheet)],
    ['first drawn', ...rows.map((r) => sec(r.drawn))],
    ['Ready', ...rows.map((r) => sec(r.ready))],
    ['  of it reading the file', ...rows.map((r) => (r.read ? sec(r.read) : '?'))],
    ['  longest pause (frame / task)', ...rows.map((r) => `${ms(r.openPause.frame)} / ${ms(r.openPause.task)}`)],
    ['largest sheet', ...rows.map((r) => r.largest)],
    ['jump to it', ...rows.map((r) => sec(r.jump))],
    [`down ${DOWN} screens (median, max)`, ...rows.map((r) => `${ms(med(r.down))}, ${ms(Math.max(...r.down))} ms`)],
    [`right ${RIGHT} screens (median, max)`, ...rows.map((r) => `${ms(med(r.right))}, ${ms(Math.max(...r.right))} ms`)],
    ['  longest pause (frame / task)', ...rows.map((r) => `${ms(r.scrollPause.frame)} / ${ms(r.scrollPause.task)}`)],
    ['edit to dependent on screen', ...rows.map((r) => (r.edit ? sec(r.edit.shown) : 'none'))],
    ['  to Ready', ...rows.map((r) => (r.edit ? sec(r.edit.ready) : '-'))],
    ['Control+Z on screen', ...rows.map((r) => (r.edit ? sec(r.edit.undo) : '-'))],
    ['Save as', ...rows.map((r) => `${sec(r.save)}, ${r.saveMb.toFixed(1)} MB`)],
    ['  formulas read back', ...rows.map((r) => `${r.savedFormulas}`)],
    ['  longest pause (frame / task)', ...rows.map((r) => `${ms(r.savePause.frame)} / ${ms(r.savePause.task)}`)],
    ['heap page / worker, open', ...rows.map((r) => `${ms(r.heapOpen.page)} / ${ms(r.heapOpen.worker)} MB`)],
    ['heap page / worker, end', ...rows.map((r) => `${ms(r.heapEnd.page)} / ${ms(r.heapEnd.worker)} MB`)],
    ['workbook language to en-US', ...rows.map((r) => sec(r.relocale))],
    ['page errors', ...rows.map((r) => `${r.errors.length}`)],
  ]
  const widths = lines[0]!.map((_, i) => Math.max(...lines.map((l) => l[i]!.length)))
  for (const l of lines) console.log(l.map((c, i) => c.padEnd(widths[i]!)).join(' | '))
}

const paths = (process.env.BIG_XLSX ?? '').split(':').filter(Boolean)
/** A published app to measure instead of the built one (`BIG_URL=https://xlcalc.pridelab.dev/demo/`): nothing is built or served. */
const live = process.env.BIG_URL
if (live && paths.length) throw new Error('BIG_URL takes the synthetic book only: books from files are not opened in a published app')
if (!live) await build({ configFile: config, logLevel: 'error' })
const server = live ? null : await preview({ configFile: config, preview: { port: 0 }, logLevel: 'error' })
const base = live ? live.replace(/\/?$/, '/') : server!.resolvedUrls!.local[0]!
const browser = await chromium.launch()
const rows: Row[] = []
try {
  const plans = paths.length ? paths.map((p) => () => planFile(p)) : [synthPlan]
  for (const [i, make] of plans.entries()) {
    const plan = await make()
    console.log(`[${plan.label}] opens ${plan.first}, largest ${plan.largest.name} (${plan.largest.content} cells with data), ${plan.edits.length} cells to edit`)
    rows.push(await check(browser, base, plan, i + 1))
  }
} finally {
  await browser.close()
  if (server) await new Promise<void>((done) => server.httpServer.close(() => done()))
}
print(rows)
writeFileSync(resolve(results, 'big-check.json'), JSON.stringify(rows, null, 2))
const bad = rows.flatMap((r) => [...r.failures, ...r.errors.map((e) => `page error: ${e}`)].map((f) => `[${r.label}] ${f}`))
console.log(`screens in ${screens}`)
if (bad.length) throw new Error(bad.join('\n'))
