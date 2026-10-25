// Screenshots of the example at phone and desktop widths: every sheet of the "all formatting" workbook and the editing states.
// Run: npm run screens (Node 22.18+ strips the types). Output: test-results/screens/, not in git.
import { mkdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { chromium, type Page } from 'playwright'
import { createServer } from 'vite'
import { quoteSheetIfNeeded } from 'xlcalc/a1'
import { longPress, longPressAt } from './touch.ts'

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

/** Underline, double underline and strikethrough must be computed on the element that holds the text (they do not reach an overflow run). */
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

/** Every button, input and tab outside the grid (within `scope`) is at least 44 × 44 px on a phone (the Name Box by its height); the desktop is unchanged. */
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
 * Excel's data validation on the "all formatting" workbook's rules sheet: the input message with the arrow, the list (Alt+Down; a tap on a phone), an item entered, a long list, and the alerts of the
 * three styles with what their buttons leave in the cell.
 */
async function validationScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const go = async (cell: string) => {
    await page.goto(`${base}?sheet=${encodeURIComponent(`'Проверка данных'!${cell}`)}`)
    await settle(page)
    // the jump's flash (2 s) ends first
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
  console.log(`validation ${size.name}: the input message, the list and an item entered; Cancel, Yes and OK left 5, 50 and ASAP`)
}

/**
 * Excel's commands in the cell menu: Freeze Panes from the menu with the sheet frozen at the cell, the Trace Precedents panel
 * for a chain of formulas with a reference in it followed to another sheet, and Reapply of the display filter after an edit
 * made a row fail its conditions (the row before and after).
 */
async function commandScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const go = async (place: string) => {
    await open(page, place)
    // the jump's flash (2 s) ends first
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
  // Keys in the menu: focus on its first enabled item (Cut is disabled on a formula), Down to the next one, shown as focused.
  await page.keyboard.press('ArrowDown')
  const focused = await page.evaluate(() => (document.activeElement as HTMLElement | null)?.dataset.xsMenu)
  if (focused !== 'paste') throw new Error(`Down in the cell menu went to ${focused}`)
  await shot(page, size.name, '50-freeze-menu-keys')
  await choose('freeze')
  await page.waitForSelector('[data-cell="2:1"].xs-fz-corner')
  await page.locator('.xs-scroller').evaluate((el) => el.scrollTo({ top: 100, left: 120 }))
  await page.waitForTimeout(400)
  await shot(page, size.name, '51-frozen')
  // The frozen corner stays in view whatever the scroll position.
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
  // The panel's own targets: the formula bar's references for a formula cell are not 44 px on a phone (FormulaBar's, not the panel's).
  await checkTargets(page, size.mobile, '[data-xs-explain]')
  const ref = page.locator('[data-xs-explain-ref="Формулы!B2"]').first()
  if (size.mobile) await ref.tap()
  else await ref.click()
  await page.waitForFunction(() => document.querySelector('[data-xs-explain-place]')?.textContent === 'Формулы!B2', null, { timeout: 5000 })
  await page.waitForTimeout(300)
  await shot(page, size.name, '53-precedents-followed')
  await page.locator('[data-xs-explain-close]').click()

  // The filter keeps «ожидает» and «оплачен»; row 3 no longer matches once its status is typed over, and stays until Reapply.
  await go('Фильтр!B3')
  const funnel = size.mobile ? page.locator('.xs-bar [data-xs-funnel]') : page.locator('.xs-head').nth(1).locator('.xs-funnel')
  await funnel.click({ force: true })
  await page.waitForSelector('[data-xs-values] .xs-value-all')
  await page.locator('.xs-value', { hasText: 'отклонён' }).click()
  await page.locator('.xs-value', { hasText: '(пусто)' }).click()
  await page.locator('[data-xs-apply]').click()
  await page.waitForSelector('[data-xs-filter-status]')
  await page.waitForFunction(() => !document.querySelector('.xs-loading'))
  // Through the formula bar: Cyrillic arrives as text input without a key press, which types over a cell only in an editor.
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
  console.log(`commands ${size.name}: freeze at B3 and unfreeze, the precedents of B12 and Формулы!B2, row 3 hidden after ${size.mobile ? 'the status line\'s Reapply' : 'Ctrl+Alt+L'}`)
}

/**
 * The formula bar's layer lines up with its input's transparent text to the pixel: the bar drawn with the layer's text, then with the input's own
 * (both black, the caret hidden), must give the same picture, except for faint antialiasing; with `refs` the layer colors them by highlights.
 */
async function checkOverlay(page: Page, label: string, refs = false): Promise<void> {
  if (refs) {
    const painted = await page.evaluate(() => [...CSS.highlights.keys()].filter((k) => k.startsWith('xs-ref-') && CSS.highlights.get(k)!.size > 0).length)
    if (!painted) throw new Error(`the formula bar's references are not colored (${label})`)
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
  throw new Error(`the formula bar's layer does not line up with its input (${label}): ${diff} pixels differ, test-results/screens/overlay-${name}-*.png`)
}

/** The formula bar with the mouse, as in Excel and a code editor: Ctrl+click (Cmd on a Mac) on a reference goes to it, a drag starts editing with the dragged-over text selected, a click moves an editor open in the cell to the formula bar at the caret. */
async function barMouse(page: Page): Promise<void> {
  const text = '=СУММ(A:A)-СУММ(A2:A5)+Форматы!C2*0'
  // the layer's characters line up with the input's: their boxes say where to press
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
  if (!moved.editing || moved.at[0] !== 3 || moved.at[1] !== 3 || (await page.locator('.xs-editor:not(.xs-editor-idle)').count())) throw new Error(`a click on the formula bar did not move the cell editor there at the caret: ${JSON.stringify(moved)}`)
  await page.keyboard.press('Escape')
  console.log(`formula bar mouse: Ctrl+click went to A2:A5, a drag selected [1, 5] in an edit, the cell editor moved to the formula bar at 3`)
}

/**
 * Going to a reference from the formula bar outside editing: with Ctrl or Cmd held (in either order) the reference under the mouse is underlined
 * (with its own highlight) and the pointer becomes a hand, both gone with the key; on a touch screen a long press on one goes to its range, a tap edits.
 */
async function barGo(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const text = '=СУММ(A:A)-СУММ(A2:A5)+Форматы!C2*0'
  const centre = (i: number) =>
    page.evaluate((i) => {
      const r = new Range()
      const t = document.querySelector('.xs-fx-layer')!.firstChild!
      r.setStart(t, i)
      r.setEnd(t, i + 1)
      const b = r.getBoundingClientRect()
      return { x: b.left + b.width / 2, y: b.top + b.height / 2 }
    }, i)
  const state = () =>
    page.evaluate(() => {
      const field = document.querySelector<HTMLTextAreaElement>('.xs-fx-input')!
      // the highlights whose rule underlines, and the text of their ranges
      const rules = [...document.adoptedStyleSheets].flatMap((s) => [...s.cssRules].map((r) => r.cssText))
      const names = rules.filter((r) => r.includes('underline')).map((r) => /::highlight\(([\w-]+)\)/.exec(r)?.[1])
      return {
        go: field.classList.contains('xs-fx-go'),
        cursor: getComputedStyle(field).cursor,
        underlined: names.flatMap((n) => [...(CSS.highlights.get(n!) ?? [])].map((r) => r.toString())),
        address: document.querySelector<HTMLInputElement>('.xs-address')!.value,
        editing: !!document.querySelector('.xs-bar-editing'),
        focused: document.activeElement === field,
      }
    })
  await open(page, 'Формулы!B6')
  if (!size.mobile) {
    const ref = await centre(text.indexOf('A2:A5') + 2)
    await page.keyboard.down('ControlOrMeta')
    await page.mouse.move(ref.x, ref.y)
    const held = await state()
    if (!held.go || held.cursor !== 'pointer' || held.underlined.join() !== 'A2:A5') throw new Error(`Cmd held over A2:A5 did not underline it with a hand: ${JSON.stringify(held)}`)
    await page.locator('.xs-bar').screenshot({ path: resolve(out, `${size.name}-63-bar-go-hover.png`) })
    await page.keyboard.up('ControlOrMeta')
    const up = await state()
    if (up.go || up.cursor === 'pointer' || up.underlined.length) throw new Error(`the released key left the underline or the hand pointer: ${JSON.stringify(up)}`)
    const off = await centre(text.indexOf('СУММ') + 1)
    await page.mouse.move(off.x, off.y)
    await page.keyboard.down('ControlOrMeta')
    const name = await state()
    await page.mouse.move(ref.x, ref.y)
    const later = await state()
    await page.keyboard.up('ControlOrMeta')
    if (name.go || name.underlined.length || !later.go || later.underlined.join() !== 'A2:A5') throw new Error(`the key first, then the mouse onto A2:A5: ${JSON.stringify({ name, later })}`)
    return void console.log(`formula bar navigation ${size.name}: Cmd held over A2:A5 underlined it with a hand pointer, releasing the key removed both, nothing over СУММ`)
  }
  // collapsed on a phone, the formula bar shows its first line: A:A
  const ref = await centre(text.indexOf('A:A') + 1)
  await longPressAt(page, ref.x, ref.y)
  await page.waitForTimeout(200)
  const went = await state()
  await shot(page, size.name, '64-bar-go-touch')
  if (went.address !== 'A:A' || went.editing || went.focused) throw new Error(`a long press on A:A did not go to it alone: ${JSON.stringify(went)}`)
  await open(page, 'Формулы!B6')
  await page.touchscreen.tap(ref.x, ref.y)
  await page.waitForTimeout(200)
  const tapped = await state()
  if (tapped.address !== 'B6' || !tapped.editing) throw new Error(`a tap on A:A did not start the edit: ${JSON.stringify(tapped)}`)
  await page.keyboard.press('Escape')
  console.log(`formula bar navigation ${size.name}: a long press on A:A went to it without an edit, a tap on it started editing`)
}

/** Excel's formula bar over a six-line cell: one line high with the grid in place, expanded by its arrow, a click's caret where it was placed, and the caret moved to the last visible line in the collapsed input. */
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
  if (collapsed.grid !== plain.grid) throw new Error(`a six-line cell moved the grid from ${plain.grid} to ${collapsed.grid} px`)
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
  // A click opens the editor in the formula bar with the caret where it was placed, as in Excel; moving the caret to the end, on the sixth line, scrolls the collapsed input and its layer there.
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
  // The open editor is as tall as the text it covers, collapsed and expanded (Ctrl+Shift+U on a desktop, the arrow on a touch screen): the grid does not move.
  const editing = await measure()
  if (size.mobile) await press('[data-xs-bar-toggle]')
  else await page.keyboard.press('Control+Shift+KeyU')
  await page.waitForSelector('.xs-bar-editing .xs-fx-open')
  const editingOpen = await measure()
  if (editing.grid !== collapsed.grid || editingOpen.grid !== expanded.grid) throw new Error(`the open editor moved the grid: ${JSON.stringify({ collapsed, editing, expanded, editingOpen })}`)
  // A long formula's draft wraps over lines with its references colored: the layer lines up with the input, expanded and collapsed.
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
  if ([plain, collapsed, expanded].some((m) => m.wide)) throw new Error(`the page scrolls horizontally: ${JSON.stringify({ plain, collapsed, expanded })}`)
  console.log(`formula bar ${size.name}: ${collapsed.bar} px collapsed over six lines (one line: ${plain.bar} px), ${expanded.bar} px expanded, the same with the editor open; the caret's line scrolled into view (${caret.top} of ${caret.room} px)`)
}

/** Merge & Center on the app's toolbar (spa.html): its dropdown open (on a phone in the toolbar's menu), then A1:B2 merged and centered. */
async function mergeScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const press = (selector: string) => (size.mobile ? page.locator(selector).tap() : page.locator(selector).click())
  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
  // through the formula bar: Cyrillic arrives as text input without a key press, which types over a cell only in an editor
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
  await page.waitForSelector('[data-tb="unmerge-cells"]')
  await shot(page, size.name, '56-merge-menu')
  await checkTargets(page, size.mobile, '.xs-toolbar')
  await press('[data-tb="merge-center"]')
  await page.waitForSelector('[data-cell="1:1"][colspan="2"][rowspan="2"]', { timeout: 30_000 })
  await page.waitForSelector('[data-tb="merge"][aria-pressed="true"]')
  // the phone's toolbar menu closed, the jump's flash (2 s) over
  if (size.mobile) await press('[data-tb="more"]')
  await page.waitForTimeout(2200)
  await shot(page, size.name, '57-merged')
  const wide = await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth)
  if (wide) throw new Error(`merge ${size.name}: the page scrolls horizontally`)
  console.log(`merge ${size.name}: the dropdown under Merge & Center, A1:B2 merged and centered, the button pressed`)
}

/**
 * Excel's typo dialog in the app (spa.html, the local source, xlcalc 0.44): `=(1+2` with Enter offers `=(1+2)`, then No
 * brings up the message about the missing parenthesis; its OK leaves the editor open with the typed text.
 */
async function typoScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const press = (selector: string) => (size.mobile ? page.locator(selector).tap() : page.locator(selector).click())
  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
  await press('[data-cell="3:2"]')
  await press('.xs-fx')
  await page.keyboard.press('ControlOrMeta+A')
  await page.keyboard.type('=(1+2')
  await page.keyboard.press('Enter')
  await page.waitForSelector('[data-xs-alert="typo"]', { timeout: 30_000 })
  await page.waitForTimeout(100)
  await shot(page, size.name, '66-typo')
  if (size.mobile) await checkTargets(page, true, '.xs-va')
  await press('[data-xs-alert-button="no"]')
  await page.waitForSelector('[data-xs-alert="parenthesis"]')
  await page.waitForTimeout(100)
  await shot(page, size.name, '67-typo-parenthesis')
  await press('[data-xs-alert-button="ok"]')
  await page.waitForSelector('[data-xs-alert]', { state: 'detached' })
  const wide = await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth)
  if (wide) throw new Error(`typo ${size.name}: the page scrolls horizontally`)
  console.log(`typo ${size.name}: =(1+2 offered =(1+2), No said the parenthesis is missing, OK left the editor open`)
}

/**
 * Excel's argument tooltip in the app (spa.html, the local source): SUMIFS at its third group of criteria in a cell near
 * the right edge (a phone's viewport is narrower than the tooltip: it wraps inside it), then ROUND inside a sum in the
 * formula bar. The tooltip stays inside the window, never together with the function list, and the page does not scroll horizontally.
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
    if (got.current !== current || !got.inside || got.list || got.wide) throw new Error(`tooltip ${size.name} ${where}: ${JSON.stringify(got)}, want ${current} inside the window, alone`)
    return got
  }
  await page.goto(`${base}spa.html?lang=ru`)
  await page.waitForSelector('[data-cell="1:1"]')
  await page.waitForFunction(() => document.querySelector('[data-spa-status]')?.textContent?.trim() === 'Готово', null, { timeout: 30_000 })
  // `=` by its key opens the cell editor; Cyrillic then arrives as text input into the editor
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
  // Excel's Cancel and Enter next to the formula bar, enabled while the edit is open
  await shot(page, size.name, '60-bar-buttons')
  await page.keyboard.press('Escape')
  console.log(`tooltip ${size.name}: SUMIFS in a cell (${cell.lines} lines), ROUND in the formula bar (${bar.lines} lines), inside the window, no list, no horizontal scroll`)

  // Excel's Insert Function: the dialog over the full list, then a category with the syntax and description of the selected one.
  await press('[data-xs-bar-fx]')
  await page.waitForSelector('[data-xs-insert-function] [data-xs-if-item]')
  await page.waitForFunction(() => document.querySelector('[data-xs-if-about]')?.textContent?.trim())
  await shot(page, size.name, '61-insert-function')
  if (size.mobile) await checkTargets(page, true, '.xs-ifn')
  await page.locator('[data-xs-if-category]').selectOption('date')
  await page.locator('[data-xs-if-item]').nth(2).click()
  await page.waitForTimeout(100)
  await shot(page, size.name, '62-insert-category')
  const wide = await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth)
  if (wide) throw new Error(`insert ${size.name}: the page scrolls horizontally`)
  await page.keyboard.press('Escape')
  await page.waitForSelector('[data-xs-insert-function]', { state: 'detached' })
}

/** Excel's Go To (F5) after a jump, with its alert for a bad reference, and the Symbol dialog (the example's Ω button): a subset, Insert into a new edit, the Special Characters tab into the same edit. */
async function dialogScreens(page: Page, size: (typeof SIZES)[number]): Promise<void> {
  const flat = async (what: string) => {
    if (await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth)) throw new Error(`${what} ${size.name}: the page scrolls horizontally`)
  }
  const draft = () => page.evaluate(() => (document.querySelector('.xs-editor') as HTMLTextAreaElement).value)
  await open(page, 'Оформление!B2')
  await page.locator('.xs-scroller').focus()
  await page.keyboard.press('F5')
  await page.locator('[data-xs-goto-ref]').fill('C5')
  await page.keyboard.press('Enter')
  await page.waitForSelector('[data-xs-goto]', { state: 'detached' })
  await page.keyboard.press('F5')
  await page.waitForSelector('[data-xs-goto-item]')
  const offered = await page.locator('[data-xs-goto-ref]').inputValue()
  if (offered !== 'Оформление!$B$2') throw new Error(`goto ${size.name}: the Reference box offers ${offered}, want Оформление!$B$2`)
  await shot(page, size.name, '70-goto')
  if (size.mobile) await checkTargets(page, true, '.xs-goto')
  await flat('goto')
  await page.locator('[data-xs-goto-ref]').fill('B2+C1')
  await page.keyboard.press('Enter')
  await page.waitForSelector('[data-xs-alert="notice"]')
  await shot(page, size.name, '71-goto-bad')
  await page.keyboard.press('Enter')
  await page.waitForSelector('[data-xs-alert]', { state: 'detached' })
  await page.keyboard.press('Escape')
  await page.waitForSelector('[data-xs-goto]', { state: 'detached' })

  await page.locator('[data-symbol]').click()
  await page.waitForSelector('[data-xs-symbol] [data-xs-sym-code]')
  const cols = await page.evaluate(() => document.querySelector('.xs-sym-row')!.children.length)
  if (!size.mobile && cols !== 16) throw new Error(`symbol ${size.name}: ${cols} characters a row, want 16`)
  await shot(page, size.name, '72-symbol')
  if (size.mobile) await checkTargets(page, true, '.xs-sym')
  await flat('symbol')
  await page.locator('[data-xs-sym-subset]').selectOption('cyrillic')
  await page.locator('[data-xs-sym-code="0416"]').click()
  await page.locator('[data-xs-sym-insert]').click()
  await shot(page, size.name, '73-symbol-cyrillic')
  await page.locator('[data-xs-sym-tab="special"]').click()
  await page.locator('[data-xs-sym-special="emDash"]').dblclick()
  await shot(page, size.name, '74-symbol-special')
  if (size.mobile) await checkTargets(page, true, '.xs-sym')
  const typed = await draft()
  if (typed !== 'Ж—') throw new Error(`symbol ${size.name}: the edit holds ${typed}, want Ж—`)
  await page.keyboard.press('Escape')
  await page.waitForSelector('[data-xs-symbol]', { state: 'detached' })
  const focused = await page.evaluate(() => document.activeElement?.classList.contains('xs-editor'))
  if (!focused) throw new Error(`symbol ${size.name}: the edit lost focus after Close`)
  await page.keyboard.press('Escape')
  console.log(`dialogs ${size.name}: Go To offered ${offered} after a jump and alerted on B2+C1; Symbol ${cols} a row, Ж and — inserted into one edit, focus back in it`)
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

  // A formula's ranges highlighted, its references in the formula bar.
  await open(page, 'Формулы!B6')
  await shot(page, size.name, '20-formula-refs')
  await checkOverlay(page, 'a formula with references', true)
  if (!size.mobile) await barMouse(page)
  await barGo(page, size)
  // A selection with its totals.
  await open(page, 'Форматы!B2:C11')
  await shot(page, size.name, '21-selection-totals')
  // A jump far down a long sheet flashes the location.
  await open(page, "'Длинный лист'!B30000")
  await page.waitForTimeout(300)
  await shot(page, size.name, '22-jump-flash')

  // Frozen panes stay while the body scrolls; an entire column of a long sheet counts only its cells.
  await open(page, 'Закрепление!M150')
  await page.waitForTimeout(2200)
  await shot(page, size.name, '28-freeze-scrolled')
  await open(page, "'Длинный лист'!C:C")
  await page.waitForTimeout(2200)
  await shot(page, size.name, '29-column-count')

  // Touch targets on a phone, with the search bar and the funnel's list open; a desktop's stay unchanged.
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

  // 0.2.0: search with its matches, a column's funnel (on a phone, full screen), a filtered and sorted sheet, a recalculation.
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

  // Excel's outline: a group's + (on a phone a tap next to it, within touch reach), then a column level button.
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
  await typoScreens(page, size)
  await tipScreens(page, size)
  await dialogScreens(page, size)

  // Editing in the cell: typing over a cell, then the pending edit.
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

  // A conflict keeps the text in the editor and explains.
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
