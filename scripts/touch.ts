// Touch input for the browser checks: Chromium's touch events through CDP, as a phone's touch screen sends them.
import type { Page } from 'playwright'

/** A finger held still on an element for 0.7 s, as Chromium's touch input receives it. */
export async function longPress(page: Page, selector: string): Promise<void> {
  const box = (await page.locator(selector).boundingBox())!
  const point = { x: box.x + box.width / 2, y: box.y + box.height / 2 }
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [point] })
  await page.waitForTimeout(700)
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] })
  await cdp.detach()
}

/** A finger held still at a point on the page for 0.7 s, like `longPress` (a point inside an element, such as a character of a text). */
export async function longPressAt(page: Page, x: number, y: number): Promise<void> {
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x, y }] })
  await page.waitForTimeout(700)
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] })
  await cdp.detach()
}

/** Two fingers around a point of the page, spread from `from` to `to` px apart in steps, as a pinch zoom. */
export async function pinch(page: Page, x: number, y: number, from: number, to: number, steps = 8): Promise<void> {
  const cdp = await page.context().newCDPSession(page)
  const at = (d: number) => [
    { x: x - d / 2, y, id: 1 },
    { x: x + d / 2, y, id: 2 },
  ]
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: at(from) })
  for (let i = 1; i <= steps; i++) {
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: at(from + ((to - from) * i) / steps) })
    await page.waitForTimeout(16)
  }
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] })
  await cdp.detach()
}
