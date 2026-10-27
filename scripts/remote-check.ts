// A workbook on a server in a browser: one localSource on this Node process, served by serveSource under /sheet/1/ next to
// the example's dev server, and three pages of example/remote.html as three viewers, two editors and a reader. Checks that an
// edit on one page appears on the others without a reload, that one editor's Undo takes back the other's edit (one history),
// that the reader's sheet is read-only and has no toolbar, and that close() reloads every page with its message.
// Run: npm run check:remote. Output: test-results/screens/remote-*.png.
import { mkdirSync } from 'node:fs'
import { createServer as createHttpServer } from 'node:http'
import { resolve } from 'node:path'
import { chromium, type Page } from 'playwright'
import { createServer } from 'vite'
import { nodeHandler, started } from './node-handler.ts'

const root = resolve(import.meta.dirname, '..')
const out = resolve(root, 'test-results/screens')
mkdirSync(out, { recursive: true })

const vite = await createServer({ configFile: resolve(root, 'vite.example.config.ts'), server: { middlewareMode: true, hmr: false }, logLevel: 'error' })
// The package's code as Vite serves it to the page, run on this process
const { localSource } = (await vite.ssrLoadModule(resolve(root, 'src/local/index.ts'))) as typeof import('../src/local')
const { serveSource } = (await vite.ssrLoadModule(resolve(root, 'src/remote/server.ts'))) as typeof import('../src/remote/server')

const source = localSource({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 'Выручка' }, B1: { value: 2 }, B2: { value: 3 }, B3: { formula: 'SUM(B1:B2)' } } }] }, { calc: 'inline' })
const serving = serveSource(source, { access: (req) => (req.headers.get('x-user')?.startsWith('editor') ? 'write' : 'read') })

// The sheet's routes go to the served source, the rest to Vite
const sheetRoutes = nodeHandler((req) => serving.handle(req))
const server = createHttpServer((req, res) => (req.url?.startsWith('/sheet/1/') ? sheetRoutes(req, res) : vite.middlewares(req, res)))
const url = await started(server)

const browser = await chromium.launch()
const errors: string[] = []
const failures: string[] = []
const check = (ok: boolean, what: string) => {
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`)
  if (!ok) failures.push(what)
}

async function viewer(user: string): Promise<Page> {
  const page = await browser.newPage({ viewport: { width: 1280, height: 720 } })
  page.on('pageerror', (e) => errors.push(`${user}: ${e.message}`))
  await page.goto(`${url}/remote.html?user=${user}`)
  await page.waitForFunction(() => document.querySelector('[data-cell="3:2"]')?.textContent?.trim() === '5', null, { timeout: 30_000 })
  return page
}
const shows = (page: Page, cell: string, text: string) =>
  page.waitForFunction(([c, t]) => document.querySelector(`[data-cell="${c}"]`)?.textContent?.trim() === t, [cell, text] as const, { timeout: 10_000 }).then(
    () => true,
    () => false,
  )
async function type(page: Page, cell: string, text: string): Promise<void> {
  await page.locator(`[data-cell="${cell}"]`).click()
  await page.keyboard.type(text)
  await page.keyboard.press('Enter')
}

try {
  const a = await viewer('editor-a')
  const b = await viewer('editor-b')
  const r = await viewer('reader')
  check((await a.locator('[data-remote-access="write"]').count()) === 1 && (await r.locator('[data-remote-access="read"]').count()) === 1, 'the server gives each viewer its access')
  check((await a.locator('.xs-toolbar').count()) === 1 && (await r.locator('.xs-toolbar').count()) === 0, "the editor has the toolbar, the reader's sheet has none")

  // Insert Function asks for every function's signature at once: one batch, not a request each
  const sent: string[] = []
  a.on('request', (req) => void (req.url().includes('/sheet/1/') && sent.push(new URL(req.url()).pathname.split('/').pop()!)))
  await a.locator('[data-cell="4:2"]').click()
  const begun = Date.now()
  await a.locator('[data-xs-bar-fx]').click()
  await a.waitForFunction(() => (document.querySelector('[data-xs-if-category]') as HTMLSelectElement | null)?.options.length! > 2, null, { timeout: 30_000 })
  const listed = Date.now() - begun
  const batches = sent.filter((x) => x === 'batch').length
  const calls = sent.filter((x) => x === 'call').length
  console.log(`Insert Function: categories in ${listed} ms, ${batches} batch and ${calls} single requests`)
  check(batches >= 1 && batches + calls <= 5, 'Insert Function reads the signatures in a batch')
  await a.keyboard.press('Escape')
  await a.keyboard.press('Escape')

  await type(a, '1:2', '10')
  check(await shows(a, '3:2', '13'), "the editor's own sheet recalculates")
  check((await shows(b, '1:2', '10')) && (await shows(b, '3:2', '13')), "the other editor's sheet shows the edit without a reload")
  check(await shows(r, '3:2', '13'), "the reader's sheet shows the edit without a reload")
  await a.screenshot({ path: resolve(out, 'remote-editor.png') })

  await b.locator('.xs-scroller').focus()
  await b.keyboard.press('ControlOrMeta+Z')
  check((await shows(a, '1:2', '2')) && (await shows(a, '3:2', '5')), "the other editor's Undo takes back the edit (one history for the workbook)")

  await type(r, '2:2', '99')
  check(!(await shows(r, '2:2', '99')) && (await shows(a, '2:2', '3')), "the reader's typing changes nothing")
  await r.screenshot({ path: resolve(out, 'remote-reader.png') })

  serving.close('Открыта новая версия книги')
  const told = await a.waitForFunction(() => document.body.innerText.includes('Открыта новая версия книги'), null, { timeout: 10_000 }).then(() => true, () => false)
  await a.waitForTimeout(500)
  await a.screenshot({ path: resolve(out, 'remote-closed.png') })
  check(told && errors.length === 0, "close() shows every page its message, without a page error")
} finally {
  console.log(`page errors: ${errors.length}${errors.length ? `\n${errors.join('\n')}` : ''}`)
  await browser.close()
  await vite.close()
  server.closeAllConnections()
  server.close()
  await source.close()
}
if (failures.length || errors.length) process.exit(1)
