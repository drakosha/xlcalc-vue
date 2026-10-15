// Screenshots of the example at a phone's and a desktop's width: every sheet of the "all formatting" book and the editing states.
// Run: npm run screens (node 22.18+ strips the types). Output: test-results/screens/, not in git.
import { mkdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { chromium, type Page } from 'playwright'
import { createServer } from 'vite'
import { quoteSheetIfNeeded } from 'xlcalc/a1'
import { longPress } from './touch.ts'

const root = resolve(import.meta.dirname, '..')
const out = resolve(root, 'test-results/screens')
mkdirSync(out, { recursive: true })

const SHEETS = ['Оформление', 'Форматы', 'Гистограммы', 'Ссылки', 'Примечания', 'Формулы', 'Объединения', 'Размеры', 'Закрепление', 'Таблица', 'Длинный лист', 'Запертые', 'Перелив', 'Плашка', 'Фильтр', 'Условное форматирование', 'Группировка', 'Проверка данных', 'Расчёт']
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

/** Every button, field and tab outside the grid (within `scope`) at least 44 × 44 px on a phone (the address box by its height); a desktop unchanged. */
async function checkTargets(page: Page, mobile: boolean, scope = '.xs-sheet'): Promise<void> {
  const boxes = await page.evaluate((scope) =>
    [...document.querySelectorAll<HTMLElement>(['button', 'input', 'textarea', '[role=tab]', 'a'].map((x) => `${scope} ${x}`).join(', '))]
      .filter((el) => !el.closest('.xs-table') && !el.matches('.xs-editor-idle') && el.getClientRects().length > 0)
      .map((el) => {
        const target = el.matches('input[type=checkbox]') ? (el.closest('label') ?? el) : el
        const r = target.getBoundingClientRect()
        return { what: `${el.tagName.toLowerCase()}.${el.className || el.getAttribute('title') || ''}`.slice(0, 60), w: Math.round(r.width), h: Math.round(r.height) }
      }),
    scope,
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

/**
 * Excel's data validation on the "all formatting" book's sheet of rules: the input message with the arrow, the list (Alt+Down; a tap on a phone), an item entered, a long list, and the alerts of the
 * three styles with what their buttons leave in the cell.
 */
async function validationScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const go = async (cell: string) => {
    await page.goto(`${base}?sheet=${encodeURIComponent(`'Проверка данных'!${cell}`)}`)
    await settle(page)
    // the jump's flash (2 s) goes first
    await page.waitForTimeout(2200)
  }
  const press = (selector: string) => (size.mobile ? page.locator(selector).tap() : page.locator(selector).click())
  const shows = (cell: string, text: string) => page.waitForFunction(([c, t]) => document.querySelector(`[data-cell="${c}"]`)?.textContent?.trim() === t, [cell, text] as const, { timeout: 5000 })
  const openList = async () => {
    await page.waitForSelector('[data-xs-dropdown-arrow]')
    if (size.mobile) await press('[data-xs-dropdown-arrow]')
    else {
      await page.locator('.xs-scroller').focus()
      await page.keyboard.press('Alt+ArrowDown')
    }
    await page.waitForSelector('[data-xs-dropdown-list]')
  }
  const enter = async (text: string, severity: string) => {
    if (size.mobile) {
      await page.locator('.xs-fx').tap()
      await page.keyboard.press('ControlOrMeta+A')
    } else await page.locator('.xs-scroller').focus()
    await page.keyboard.type(text)
    await page.keyboard.press('Enter')
    await page.waitForSelector(`[data-xs-alert="${severity}"]`)
    await page.waitForTimeout(100)
  }

  await go('B2')
  await page.waitForSelector('[data-xs-prompt]')
  await shot(page, size.name, '40-validation-prompt')
  await openList()
  await shot(page, size.name, '41-validation-list')
  if (size.mobile) await checkTargets(page, true)
  await press('[data-xs-dropdown-item="2"]')
  await shows('2:2', 'закрыт')
  await go('B8')
  await openList()
  await shot(page, size.name, '42-validation-list-long')

  await go('B4')
  await enter('500', 'stop')
  await shot(page, size.name, '43-validation-stop')
  await checkTargets(page, size.mobile)
  await press('[data-xs-alert-button="cancel"]')
  await shows('4:2', '5')
  await go('B5')
  await enter('50', 'warning')
  await shot(page, size.name, '44-validation-warning')
  await press('[data-xs-alert-button="yes"]')
  await shows('5:2', '50')
  await go('B6')
  await enter('ASAP', 'information')
  await shot(page, size.name, '45-validation-information')
  await press('[data-xs-alert-button="ok"]')
  await shows('6:2', 'ASAP')
  await go('B7')
  await enter('9', 'stop')
  await shot(page, size.name, '46-validation-stop-default')
  await press('[data-xs-alert-button="retry"]')
  console.log(`validation ${size.name}: the message, the list and an item entered, Cancel, Yes and OK left 5, 50 and ASAP`)
}

/**
 * Excel's commands on the cells' menu: Freeze Panes from the menu and the sheet frozen at the cell, the precedents' panel
 * of a chain of formulas and a reference in it followed to another sheet, Reapply of the display filter after an edit
 * took a row out of its conditions (the row before and after).
 */
async function commandScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const go = async (place: string) => {
    await open(page, place)
    // the jump's flash (2 s) goes first
    await page.waitForTimeout(2200)
  }
  const menuOn = async (cell: string) => {
    if (size.mobile) await longPress(page, `[data-cell="${cell}"]`)
    else await page.locator(`[data-cell="${cell}"]`).click({ button: 'right' })
    await page.waitForSelector('[data-xs-menu-box]')
  }
  const choose = (id: string) => (size.mobile ? page.locator(`[data-xs-menu="${id}"]`).tap() : page.locator(`[data-xs-menu="${id}"]`).click())

  await go('Форматы!B3')
  await menuOn('3:2')
  await shot(page, size.name, '50-freeze-menu')
  if (size.mobile) await checkTargets(page, true)
  // Keys in the menu: the focus on its first enabled item (Cut is off on a formula), Down to the next one, shown as focused.
  await page.keyboard.press('ArrowDown')
  const focused = await page.evaluate(() => (document.activeElement as HTMLElement | null)?.dataset.xsMenu)
  if (focused !== 'paste') throw new Error(`Down in the cells' menu went to ${focused}`)
  await shot(page, size.name, '50-freeze-menu-keys')
  await choose('freeze')
  await page.waitForSelector('[data-cell="2:1"].xs-fz-corner')
  await page.locator('.xs-scroller').evaluate((el) => el.scrollTo({ top: 100, left: 120 }))
  await page.waitForTimeout(400)
  await shot(page, size.name, '51-frozen')
  // The frozen corner stays in view whatever the scroll.
  await menuOn('2:1')
  const unfreeze = await page.locator('[data-xs-menu="unfreeze"]').textContent()
  if (unfreeze?.trim() !== 'Снять закрепление областей') throw new Error(`the frozen sheet's menu offers ${unfreeze}`)
  await choose('unfreeze')
  await page.waitForFunction(() => !document.querySelector('.xs-fz-corner, .xs-fz-row, .xs-fz-col'))

  await go('Расчёт!B12')
  await menuOn('12:2')
  await choose('precedents')
  await page.waitForSelector('[data-xs-explain-line]')
  await page.waitForTimeout(200)
  await shot(page, size.name, '52-precedents')
  // The panel's own targets: the formula bar's references of a formula cell are not 44 px on a phone (FormulaBar's, not the panel's).
  await checkTargets(page, size.mobile, '[data-xs-explain]')
  const ref = page.locator('[data-xs-explain-ref="Формулы!B2"]').first()
  if (size.mobile) await ref.tap()
  else await ref.click()
  await page.waitForFunction(() => document.querySelector('[data-xs-explain-place]')?.textContent === 'Формулы!B2', null, { timeout: 5000 })
  await page.waitForTimeout(300)
  await shot(page, size.name, '53-precedents-followed')
  await page.locator('[data-xs-explain-close]').click()

  // The filter keeps «ожидает» and «оплачен»; row 3 goes out of it when its status is typed over, and stays until Reapply.
  await go('Фильтр!B3')
  const funnel = size.mobile ? page.locator('.xs-bar [data-xs-funnel]') : page.locator('.xs-head').nth(1).locator('.xs-funnel')
  await funnel.click({ force: true })
  await page.waitForSelector('[data-xs-values] .xs-value-all')
  await page.locator('.xs-value', { hasText: 'отклонён' }).click()
  await page.locator('.xs-value', { hasText: '(пусто)' }).click()
  await page.locator('[data-xs-apply]').click()
  await page.waitForSelector('[data-xs-filter-status]')
  await page.waitForFunction(() => !document.querySelector('.xs-loading'))
  // Through the formula bar: Cyrillic comes as text input without a key press, which types over a cell only in a field.
  if (size.mobile) {
    await page.locator('[data-cell="3:2"]').tap()
    await page.locator('.xs-fx').tap()
  } else {
    await page.locator('[data-cell="3:2"]').click()
    await page.locator('.xs-fx').click()
  }
  await page.keyboard.press('ControlOrMeta+A')
  await page.keyboard.type('отклонён')
  await page.keyboard.press('Enter')
  await page.waitForFunction(() => document.querySelector('[data-cell="3:2"]')?.textContent?.trim() === 'отклонён', null, { timeout: 5000 })
  await page.waitForTimeout(300)
  await shot(page, size.name, '54-reapply-before')
  if (size.mobile) await page.locator('[data-xs-filter-reapply]').tap()
  else {
    await page.locator('.xs-scroller').focus()
    await page.keyboard.press('Control+Alt+KeyL')
  }
  await page.waitForFunction(() => !document.querySelector('[data-cell="3:2"]') && !!document.querySelector('[data-cell="4:2"]'), null, { timeout: 5000 })
  await page.waitForTimeout(300)
  await shot(page, size.name, '55-reapply-after')
  if (size.mobile) await checkTargets(page, true)
  await page.locator('[data-xs-filter-clear]').click()
  console.log(`commands ${size.name}: freeze at B3 and unfreeze, the precedents of B12 and Формулы!B2, row 3 gone after ${size.mobile ? 'the line\'s Reapply' : 'Control+Alt+L'}`)
}

/**
 * The formula bar's layer lies on its field's transparent text to the pixel: the bar drawn with the layer's text, then with the field's own
 * (both black, the caret hidden), must be the same picture, but for a faint step of antialiasing; with `refs` the layer colours them by highlights.
 */
async function checkOverlay(page: Page, label: string, refs = false): Promise<void> {
  if (refs) {
    const painted = await page.evaluate(() => [...CSS.highlights.keys()].filter((k) => k.startsWith('xs-ref-') && CSS.highlights.get(k)!.size > 0).length)
    if (!painted) throw new Error(`the bar's references are not coloured (${label})`)
  }
  const fx = page.locator('.xs-fx')
  const black = Array.from({ length: 24 }, (_, i) => `.xs-fx-layer::highlight(xs-ref-${i}) { color: #000 !important }`).join(' ')
  const style = await page.addStyleTag({ content: `.xs-fx-input { caret-color: transparent !important } .xs-fx-layer { color: #000 !important } ${black}` })
  const layer = (await fx.screenshot()).toString('base64')
  await style.evaluate((el) => (el.textContent = '.xs-fx-input { caret-color: transparent !important; color: #000 !important } .xs-fx-layer { visibility: hidden !important }'))
  const field = (await fx.screenshot()).toString('base64')
  await style.evaluate((el) => (el as HTMLElement).remove())
  const diff = await page.evaluate(async ([a, b]) => {
    const load = (src: string) =>
      new Promise<HTMLImageElement>((ok, fail) => {
        const img = new Image()
        img.onload = () => ok(img)
        img.onerror = fail
        img.src = `data:image/png;base64,${src}`
      })
    const [x, y] = await Promise.all([load(a), load(b)])
    const pixels = (img: HTMLImageElement) => {
      const c = document.createElement('canvas')
      c.width = img.width
      c.height = img.height
      const g = c.getContext('2d')!
      g.drawImage(img, 0, 0)
      return g.getImageData(0, 0, c.width, c.height).data
    }
    if (x.width !== y.width || x.height !== y.height) return -1
    const p = pixels(x)
    const q = pixels(y)
    let n = 0
    for (let i = 0; i < p.length; i += 4) if (Math.max(Math.abs(p[i]! - q[i]!), Math.abs(p[i + 1]! - q[i + 1]!), Math.abs(p[i + 2]! - q[i + 2]!)) > 4) n++
    return n
  }, [layer, field] as const)
  if (diff === 0) return
  const name = label.replace(/\W+/g, '-')
  writeFileSync(resolve(out, `overlay-${name}-layer.png`), Buffer.from(layer, 'base64'))
  writeFileSync(resolve(out, `overlay-${name}-field.png`), Buffer.from(field, 'base64'))
  throw new Error(`the bar's layer does not lie on its field (${label}): ${diff} pixels differ, test-results/screens/overlay-${name}-*.png`)
}

/** The formula bar by the mouse, as Excel's and a code editor's: Ctrl+click (Cmd on a Mac) on a reference goes to it, a drag starts the edit with the text dragged over, a click moves a field open in the cell to the bar at the caret. */
async function barMouse(page: Page): Promise<void> {
  const text = '=СУММ(A:A)-СУММ(A2:A5)+Форматы!C2*0'
  // the layer's characters lie on the field's: their boxes say where to press
  const edge = (i: number, side: 'left' | 'right') =>
    page.evaluate(
      ([i, side]) => {
        const r = new Range()
        const t = document.querySelector('.xs-fx-layer')!.firstChild!
        r.setStart(t, i)
        r.setEnd(t, i + 1)
        const b = r.getBoundingClientRect()
        return { x: side === 'left' ? b.left + 1 : b.right - 1, y: b.top + b.height / 2 }
      },
      [i, side] as const,
    )
  const field = () => page.evaluate(() => ({ editing: !!document.querySelector('.xs-bar-editing'), at: [document.querySelector<HTMLTextAreaElement>('.xs-fx-input')!.selectionStart, document.querySelector<HTMLTextAreaElement>('.xs-fx-input')!.selectionEnd], address: document.querySelector<HTMLInputElement>('.xs-address')!.value }))
  await open(page, 'Формулы!B6')
  const ref = await edge(text.indexOf('A2:A5') + 2, 'left')
  await page.keyboard.down('ControlOrMeta')
  await page.mouse.click(ref.x, ref.y)
  await page.keyboard.up('ControlOrMeta')
  await page.waitForFunction(() => document.querySelector<HTMLInputElement>('.xs-address')!.value === 'A2:A5', null, { timeout: 5000 })
  const went = await field()
  if (went.editing) throw new Error(`Ctrl+click on a reference started an edit: ${JSON.stringify(went)}`)
  await open(page, 'Формулы!B6')
  const from = await edge(1, 'left')
  const to = await edge(4, 'right')
  await page.mouse.move(from.x, from.y)
  await page.mouse.down()
  await page.mouse.move(to.x, to.y, { steps: 5 })
  await page.mouse.up()
  await page.waitForTimeout(100)
  const dragged = await field()
  if (!dragged.editing || dragged.at[0] !== 1 || dragged.at[1] !== 5) throw new Error(`a drag over СУММ did not start the edit with it selected: ${JSON.stringify(dragged)}`)
  await page.keyboard.press('Escape')
  await page.locator('.xs-scroller').focus()
  await page.keyboard.press('F2')
  await page.waitForSelector('.xs-editor:not(.xs-editor-idle)')
  const caret = await edge(3, 'left')
  await page.mouse.click(caret.x, caret.y)
  await page.waitForTimeout(100)
  const moved = await field()
  if (!moved.editing || moved.at[0] !== 3 || moved.at[1] !== 3 || (await page.locator('.xs-editor:not(.xs-editor-idle)').count())) throw new Error(`a click on the bar did not take the cell's field there at the caret: ${JSON.stringify(moved)}`)
  await page.keyboard.press('Escape')
  console.log(`bar mouse: Ctrl+click went to A2:A5, a drag selected [1, 5] in an edit, the cell's field moved to the bar at 3`)
}

/** Excel's formula bar over a cell of six lines: one line high and the grid in place, expanded by its arrow, a click's caret where it was put, and the caret moved to the last line seen in the collapsed field. */
async function barScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const press = (selector: string) => (size.mobile ? page.locator(selector).tap() : page.locator(selector).click())
  const measure = () =>
    page.evaluate(() => ({
      bar: Math.round(document.querySelector('.xs-bar')!.getBoundingClientRect().height),
      grid: Math.round(document.querySelector('.xs-scroller')!.getBoundingClientRect().top),
      wide: document.documentElement.scrollWidth > document.documentElement.clientWidth,
    }))
  await open(page, 'Оформление!A1')
  const plain = await measure()
  await open(page, 'Оформление!H8')
  await page.waitForTimeout(2200)
  const collapsed = await measure()
  if (collapsed.grid !== plain.grid) throw new Error(`a cell of six lines moved the grid from ${plain.grid} to ${collapsed.grid} px`)
  await shot(page, size.name, '36-bar-collapsed')
  await checkOverlay(page, 'six lines collapsed')
  if (size.mobile) await checkTargets(page, true)
  await press('[data-xs-bar-toggle]')
  await page.waitForSelector('.xs-fx-open')
  const expanded = await measure()
  await shot(page, size.name, '37-bar-expanded')
  await checkOverlay(page, 'six lines expanded')
  await press('[data-xs-bar-toggle]')
  await page.waitForSelector('.xs-fx:not(.xs-fx-open)')
  // A click opens the field in the bar with the caret where it was put, as Excel's; the caret taken to the end, on the sixth line, scrolls the collapsed field and its layer there.
  await press('.xs-fx')
  await page.waitForSelector('.xs-bar-editing')
  await page.waitForTimeout(100)
  const field = () =>
    page.evaluate(() => {
      const el = document.querySelector<HTMLTextAreaElement>('.xs-fx-input')!
      return { at: el.selectionStart, length: el.value.length, top: el.scrollTop, room: el.scrollHeight - el.clientHeight, layer: document.querySelector('.xs-fx-layer')!.scrollTop }
    })
  const clicked = await field()
  if (clicked.at === clicked.length) throw new Error(`a click put the caret at the end: ${JSON.stringify(clicked)}`)
  await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowDown' : 'Control+End')
  await page.waitForTimeout(100)
  const caret = await field()
  if (caret.at !== caret.length || caret.room <= 0 || caret.top < caret.room - 2 || caret.layer !== caret.top) throw new Error(`the caret's line is not in view: ${JSON.stringify(caret)}`)
  await shot(page, size.name, '38-bar-caret')
  await checkOverlay(page, 'the caret on the sixth line')
  // The open field is as high as the text it covers, collapsed and expanded (Control+Shift+U on a desktop, the arrow under a finger): the grid stays.
  const editing = await measure()
  if (size.mobile) await press('[data-xs-bar-toggle]')
  else await page.keyboard.press('Control+Shift+KeyU')
  await page.waitForSelector('.xs-bar-editing .xs-fx-open')
  const editingOpen = await measure()
  if (editing.grid !== collapsed.grid || editingOpen.grid !== expanded.grid) throw new Error(`the open field moved the grid: ${JSON.stringify({ collapsed, editing, expanded, editingOpen })}`)
  // A long formula's draft wraps over lines with its references coloured: the layer lies on the field expanded and collapsed.
  await page.keyboard.press('ControlOrMeta+A')
  await page.keyboard.insertText(`=${Array.from({ length: 12 }, (_, i) => `СУММ(A${i + 1}:C${i + 9})*$B$${i + 2}`).join('+')}`)
  await page.waitForTimeout(100)
  await checkOverlay(page, 'a long draft expanded', true)
  await shot(page, size.name, '39-bar-draft')
  if (size.mobile) await press('[data-xs-bar-toggle]')
  else await page.keyboard.press('Control+Shift+KeyU')
  await page.waitForSelector('.xs-bar-editing .xs-fx:not(.xs-fx-open)')
  await page.waitForTimeout(100)
  await checkOverlay(page, 'a long draft collapsed', true)
  await page.keyboard.press('Escape')
  if ([plain, collapsed, expanded].some((m) => m.wide)) throw new Error(`the page scrolls sideways: ${JSON.stringify({ plain, collapsed, expanded })}`)
  console.log(`bar ${size.name}: ${collapsed.bar} px collapsed over six lines (one line: ${plain.bar} px), ${expanded.bar} px expanded, the same with the field open; the caret's line scrolled into view (${caret.top} of ${caret.room} px)`)
}

/** Merge & Center on the app's toolbar (spa.html): its list open (on a phone in the toolbar's menu), then A1:B2 merged and centred. */
async function mergeScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const press = (selector: string) => (size.mobile ? page.locator(selector).tap() : page.locator(selector).click())
  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
  // through the formula bar: Cyrillic comes as text input without a key press, which types over a cell only in a field
  await press('[data-cell="1:1"]')
  await press('.xs-fx')
  await page.keyboard.type('Итого за квартал')
  await page.keyboard.press('Enter')
  await page.waitForFunction(() => document.querySelector('[data-cell="1:1"]')?.textContent?.trim() === 'Итого за квартал', null, { timeout: 30_000 })
  await press('.xs-address')
  await page.locator('.xs-address').fill('A1:B2')
  await page.keyboard.press('Enter')
  if (size.mobile) await press('[data-tb="more"]')
  await press('[data-tb="merge-menu"]')
  await page.waitForSelector('[data-tb="unmerge"]')
  await shot(page, size.name, '56-merge-menu')
  await checkTargets(page, size.mobile, '.xs-toolbar')
  await press('[data-tb="merge-center"]')
  await page.waitForSelector('[data-cell="1:1"][colspan="2"][rowspan="2"]', { timeout: 30_000 })
  await page.waitForSelector('[data-tb="merge"][aria-pressed="true"]')
  // the phone's toolbar menu closed, the jump's flash (2 s) gone
  if (size.mobile) await press('[data-tb="more"]')
  await page.waitForTimeout(2200)
  await shot(page, size.name, '57-merged')
  const wide = await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth)
  if (wide) throw new Error(`merge ${size.name}: the page scrolls sideways`)
  console.log(`merge ${size.name}: the list under Merge & Center, A1:B2 merged and centred, the button pressed`)
}

/**
 * Excel's tip with a function's arguments in the app (spa.html, the local source): SUMIFS at its third group of criteria in
 * a cell near the right edge (a phone's view is narrower than the tip: it wraps inside it), then ROUND inside a sum in
 * the formula bar. The tip stays in the window, never with the list of functions, and the page does not scroll sideways.
 */
async function tipScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const press = (selector: string) => (size.mobile ? page.locator(selector).tap() : page.locator(selector).click())
  const check = async (where: string, current: string) => {
    await page.waitForSelector('[data-xs-tip] .xs-ft-on')
    const got = await page.evaluate(() => {
      const tip = document.querySelector('[data-xs-tip]')!
      const r = tip.getBoundingClientRect()
      return {
        current: tip.querySelector('.xs-ft-on')!.textContent!.trim(),
        inside: r.left >= 0 && r.right <= document.documentElement.clientWidth && r.bottom <= window.innerHeight,
        list: !!document.querySelector('[data-xs-functions]'),
        wide: document.documentElement.scrollWidth > document.documentElement.clientWidth,
        lines: Math.round(r.height / 18),
      }
    })
    if (got.current !== current || !got.inside || got.list || got.wide) throw new Error(`tip ${size.name} ${where}: ${JSON.stringify(got)}, want ${current} in the window alone`)
    return got
  }
  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
  // `=` opens the cell's field by its key; Cyrillic then comes as text input into the field
  await press(size.mobile ? '[data-cell="1:4"]' : '[data-cell="2:6"]')
  await page.keyboard.press('=')
  await page.keyboard.type('СУММЕСЛИМН(A1:A9;B1:B9;"x";C1:C9;"y";')
  const cell = await check('in the cell', 'диапазон_условия3')
  await shot(page, size.name, '58-tip-cell')
  await page.keyboard.press('Escape')
  await press('[data-cell="3:2"]')
  await press('.xs-fx')
  await page.keyboard.type('=СУММ(1;ОКРУГЛ(2,5;')
  const bar = await check('in the formula bar', 'разряды')
  await shot(page, size.name, '59-tip-bar')
  await checkOverlay(page, 'a draft under the tip')
  await page.keyboard.press('Escape')
  console.log(`tip ${size.name}: SUMIFS in a cell (${cell.lines} lines), ROUND in the bar (${bar.lines} lines), inside the window, no list, no side scroll`)
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
  await checkOverlay(page, 'a formula with references', true)
  if (!size.mobile) await barMouse(page)
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

  // Excel's outline: a group's + (on a phone a tap beside it, within a finger's reach), then a columns' level button.
  await open(page, 'Группировка!A1')
  const plus = (await page.locator('.xs-ol-rows [data-xs-outline="6:9:2"]').boundingBox())!
  if (size.mobile) await page.touchscreen.tap(plus.x + plus.width / 2 + 4, plus.y + plus.height / 2 + 15)
  else await page.mouse.click(plus.x + plus.width / 2, plus.y + plus.height / 2)
  await page.waitForFunction(() => [...document.querySelectorAll('.xs-num')].some((n) => n.textContent!.trim() === '7'), null, { timeout: 5000 })
  await page.waitForTimeout(200)
  await shot(page, size.name, '34-outline-expanded')
  const level = page.locator('.xs-ol-levels-columns [data-xs-outline-level="3"]')
  if (size.mobile) await level.tap()
  else await level.click()
  await page.waitForFunction(() => [...document.querySelectorAll('.xs-head')].some((h) => h.textContent!.trim() === 'G'), null, { timeout: 5000 })
  await page.waitForTimeout(200)
  await shot(page, size.name, '35-outline-level')
  console.log(`outline ${size.name}: + ${size.mobile ? 'tapped 15 px below' : 'clicked'}, level 3 of the columns`)

  await barScreens(page, size)
  await validationScreens(page, size)
  await commandScreens(page, size)
  await mergeScreens(page, size)
  await tipScreens(page, size)

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
