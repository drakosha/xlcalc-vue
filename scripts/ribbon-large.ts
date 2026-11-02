// The toolbar's large buttons with rows: 2 in the app (spa.html), at 1920, 1440 and 1280 px in English and in Russian: a
// PivotTable inserted from the toolbar, its cell active so that the PivotTable Analyze and Design groups show, then every icon
// at one size, each large button's label under its icon on one or two lines and not cut, and the row's groups within the
// toolbar, what does not fit in More without labels. Screenshots go to test-results/screens/ribbon-large-<width>-<lang>.png (and
// -more.png with More open); called from screens.ts.
import { resolve } from 'node:path'
import type { Browser, Page } from 'playwright'

const WIDTHS = [1920, 1440, 1280]
const READY = ['Ready', 'Готово']
const DATA: [string, string][] = [
  ['1:1', 'Region'],
  ['1:2', 'Amount'],
  ['2:1', 'North'],
  ['2:2', '10'],
  ['3:1', 'South'],
  ['3:2', '20'],
]

const ready = (page: Page) => page.waitForFunction((ok) => ok.includes(document.querySelector('[data-spa-status]')?.textContent?.trim() ?? ''), READY, { timeout: 30_000 })

/** What the toolbar draws: each icon's size by the button that holds it, each large button with its label's lines, the row against the toolbar. */
function measure() {
  const box = (el: Element) => el.getBoundingClientRect()
  const name = (b: HTMLElement) => b.dataset.tb ?? b.dataset.tbHost ?? `${b.dataset.tbHostMenu}-menu`
  const row = document.querySelector<HTMLElement>('.xs-tb-row')!
  const icons = [...row.querySelectorAll<HTMLElement>('.xs-tool')]
    .filter((b) => !b.matches('.xs-tb-caret') && b.getClientRects().length > 0)
    .flatMap((b) => {
      const icon = b.querySelector(':scope > .xs-icon')
      return icon ? [{ button: name(b), w: Math.round(box(icon).width), h: Math.round(box(icon).height) }] : []
    })
  const large = [...row.querySelectorAll<HTMLElement>('.xs-tb-label')].map((label) => {
    const b = label.closest<HTMLElement>('.xs-tool')!
    const cell = label.closest('.xs-tb-split') ?? b
    const icon = cell.querySelector('.xs-icon')!
    const lines = [...label.querySelectorAll<HTMLElement>('.xs-tb-label-line')]
    return {
      button: name(b),
      lines: lines.map((l) => l.textContent!.trim()),
      cut: lines.some((l) => l.scrollWidth > l.clientWidth + 1),
      under: box(label).top >= box(icon).bottom - 1,
      w: Math.round(box(cell).width),
      h: Math.round(box(cell).height),
    }
  })
  const bar = box(document.querySelector('.xs-toolbar')!)
  const groups = [...row.querySelectorAll<HTMLElement>(':scope > [data-tb-block]')]
  const out = groups.filter((g) => box(g).right > bar.right + 0.5 || box(g).bottom > bar.bottom + 0.5).map((g) => g.dataset.tbBlock)
  const more = [...document.querySelectorAll<HTMLElement>('[data-tb-menu] > [data-tb-block]')].map((g) => g.dataset.tbBlock)
  // A large button's labelled button: the button itself, or a split button's arrow
  const bare = [...row.querySelectorAll<HTMLElement>(':is(.xs-tool.xs-tb-large, .xs-tb-large .xs-tool)')].filter((b) => !b.matches(':has(+ .xs-tb-caret)') && !b.querySelector('.xs-tb-label')).map(name)
  return { icons, large, bare, out, kept: groups.map((g) => g.dataset.tbBlock), more, bar: Math.round(bar.height), tool: Math.round(box(row.querySelector('[data-tb="bold"]')!).height) }
}

export async function ribbonLargeScreens(browser: Browser, base: string, out: string): Promise<void> {
  for (const width of WIDTHS) {
    for (const lang of ['en', 'ru']) {
      const at = (what: string) => `ribbon large ${width} ${lang}: ${what}`
      const context = await browser.newContext({ viewport: { width, height: 900 } })
      const page = await context.newPage()
      const errors: string[] = []
      page.on('pageerror', (e) => errors.push(e.message))
      await page.goto(`${base}spa.html?lang=${lang}&rows=2`)
      await page.waitForSelector('[data-cell="1:1"]')
      await ready(page)
      for (const [cell, text] of DATA) {
        await page.locator(`[data-cell="${cell}"]`).click()
        await page.keyboard.type(text)
        await page.keyboard.press('Enter')
      }
      await ready(page)
      await page.locator('[data-cell="2:1"]').click()
      if (await page.locator('[data-tb-menu] [data-tb="pivot-table"]').count()) await page.locator('[data-tb="more"]').click()
      await page.locator('[data-tb="pivot-table"]').click()
      await page.locator('[data-xs-pivot-dialog] [data-xs-dialog-ok]').click()
      await page.waitForSelector('[data-xs-pf]', { timeout: 30_000 })
      await page.locator('[data-xs-pf-field="0"] input').click()
      await page.locator('[data-xs-pf-field="1"] input').click()
      await page.waitForFunction(() => document.querySelector('[data-cell="4:2"]')?.textContent?.trim() === '10', null, { timeout: 30_000 })
      await page.locator('[data-cell="4:1"]').click()
      await page.waitForSelector('[data-tb="pivot-subtotals"]', { state: 'attached' })
      await page.mouse.move(width / 2, 880)
      await page.waitForTimeout(300)
      const got = await page.evaluate(measure)
      await page.screenshot({ path: resolve(out, `ribbon-large-${width}-${lang}.png`), clip: { x: 0, y: 0, width, height: 200 } })
      const sizes = [...new Set(got.icons.map((i) => `${i.w}x${i.h}`))]
      if (sizes.length !== 1) throw new Error(at(`icons of ${sizes.length} sizes: ${JSON.stringify(got.icons.filter((i) => `${i.w}x${i.h}` !== sizes[0]))} besides ${sizes[0]}`))
      const bad = got.large.filter((l) => l.cut || !l.under || !l.lines[0])
      if (bad.length) throw new Error(at(`large buttons cut or without the label under the icon: ${JSON.stringify(bad)}`))
      if (got.out.length) throw new Error(at(`groups outside the toolbar: ${got.out.join(', ')}`))
      if (got.bare.length) throw new Error(at(`large buttons without a label: ${got.bare.join(', ')}`))
      if (errors.length) throw new Error(at(`page errors: ${errors.join('; ')}`))
      if (got.more.length) {
        // What does not fit is in More, drawn as on one line: without labels
        await page.locator('[data-tb="more"]').click()
        await page.mouse.move(width / 2, 880)
        await page.screenshot({ path: resolve(out, `ribbon-large-${width}-${lang}-more.png`), clip: { x: 0, y: 0, width, height: 260 } })
        const inMore = await page.locator('[data-tb-menu] .xs-tb-label').count()
        if (inMore) throw new Error(at(`${inMore} labels in More`))
      }
      const labels = got.large.map((l) => `${l.button} "${l.lines.join(' / ')}" ${l.w}x${l.h}`).join('; ')
      console.log(`${at('')}icons ${sizes[0]} (${got.icons.length}), toolbar ${got.bar} px, a small button ${got.tool} px tall; large: ${labels}; row ${got.kept.join(' ')}${got.more.length ? `; in More ${got.more.join(' ')}` : ''}`)
      await context.close()
    }
  }
}
