// The built main entry carries no engine: xlcalc/a1 and xlcalc/text are bundled in, ru-RU's formula text comes as a chunk of its
// own for a book in ru-RU, xlcalc/entry as one for a source without `editText`, and nothing is imported but vue.
import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { build } from 'vite'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

const root = resolve(import.meta.dirname, '..')
let out = ''
let js = ''
let chunks: string[] = []
const chunk = (name: string) => chunks.find((f) => f.startsWith(`${name}-`))!

beforeAll(async () => {
  out = mkdtempSync(join(tmpdir(), 'xs-dist-'))
  // A production build, as `npm run build` makes it: under vitest NODE_ENV is `test`, and a dev build names source paths.
  const env = process.env.NODE_ENV
  process.env.NODE_ENV = 'production'
  try {
    await build({ configFile: resolve(root, 'vite.config.ts'), root, mode: 'production', logLevel: 'silent', build: { outDir: out, emptyOutDir: true } })
  } finally {
    process.env.NODE_ENV = env
  }
  js = readFileSync(join(out, 'index.js'), 'utf8')
  chunks = readdirSync(out).filter((f) => f.endsWith('.js') && f !== 'index.js')
}, 60_000)
afterAll(() => rmSync(out, { recursive: true, force: true }))

const staticImports = (code: string) => [...code.matchAll(/(?:^|[;\n])\s*import\s*(?:[^'";]*?from\s*)?["']([^"']+)["']/g)].map((m) => m[1])
const dynamicImports = (code: string) => [...code.matchAll(/import\(\s*["']([^"']+)["']\s*\)/g)].map((m) => m[1])
/** The code without the bundler's comments that name the modules it took (`//#region node_modules/xlcalc/dist/a1.js`). */
const code = (text: string) => text.replace(/^\/\/#(?:end)?region.*$/gm, '')

describe('dist/index.js', () => {
  it('imports vue alone, and the ru-RU and xlcalc/entry chunks on demand', () => {
    expect(new Set(staticImports(js))).toEqual(new Set(['vue']))
    expect(chunks.sort()).toEqual([expect.stringMatching(/^ru-RU-[\w-]+\.js$/), expect.stringMatching(/^typed-entry-[\w-]+\.js$/)])
    expect(new Set(dynamicImports(js))).toEqual(new Set([`./${chunk('ru-RU')}`, `./${chunk('typed-entry')}`]))
  })
  it('carries no engine: xlcalc/a1 and xlcalc/text are a few dozen kilobytes', () => {
    expect(js.length).toBeGreaterThan(10_000)
    // the component with its dialogs and bars is some 330 KB unminified; the engine would be megabytes
    expect(js.length).toBeLessThan(400_000)
    expect(js).toContain('//#region node_modules/xlcalc/dist/a1.js')
    expect(js).toContain('//#region node_modules/xlcalc/dist/scan.js')
  })
  it('has no string xlcalc in its code, which a host may look for as the engine\'s sign', () => {
    expect(code(js)).not.toMatch(/xlcalc/)
  })
  it('comes with its style sheet', () => {
    expect(readFileSync(join(out, 'style.css'), 'utf8')).toContain('.xs-sheet')
  })
})

describe('the ru-RU chunk', () => {
  it("is ru-RU's formula text alone, without an import or the string xlcalc in its code", () => {
    const ru = readFileSync(join(out, chunk('ru-RU')), 'utf8')
    expect(ru).toContain('СУММ')
    expect(js).not.toContain('СУММ')
    expect(staticImports(ru)).toEqual([])
    expect(dynamicImports(ru)).toEqual([])
    expect(code(ru)).not.toMatch(/xlcalc/)
    // Since xlcalc 0.42 the locale carries each function's description in Russian, for Insert Function.
    expect(ru.length).toBeLessThan(75_000)
  })
})

describe('the xlcalc/entry chunk', () => {
  it('is bundled apart, whole: no import, so the main entry loads nothing more for it, and no string xlcalc in its code', () => {
    const entry = readFileSync(join(out, chunk('typed-entry')), 'utf8')
    expect(entry).toMatch(/export \{[^}]*\bentryText\b[^}]*\}/)
    expect(staticImports(entry)).toEqual([])
    expect(dynamicImports(entry)).toEqual([])
    expect(code(entry)).not.toMatch(/xlcalc/)
    // xlcalc's typed entry without the engine, the parts it shares with xlcalc/a1 and xlcalc/text included
    expect(entry.length).toBeLessThan(150_000)
  })
})
