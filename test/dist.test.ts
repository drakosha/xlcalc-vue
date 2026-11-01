// The built main entry point includes no engine: xlcalc/a1 and xlcalc/text are bundled in, ru-RU's formula text comes as a
// separate chunk for a workbook in ru-RU, xlcalc/entry as one for a source without `editText`, the PivotTable, slicer and
// timeline parts as one for a sheet with them, and nothing is imported but vue.
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
  // A production build, as `npm run build` produces it: under vitest NODE_ENV is `test`, and a dev build names source paths.
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
/** The code without the bundler's comments that name the modules it included (`//#region node_modules/xlcalc/dist/a1.js`). */
const code = (text: string) => text.replace(/^\/\/#(?:end)?region.*$/gm, '')

describe('dist/index.js', () => {
  it('imports only vue, and the ru-RU, xlcalc/entry and PivotTable chunks on demand', () => {
    expect(new Set(staticImports(js))).toEqual(new Set(['vue']))
    expect(chunks.sort()).toEqual([expect.stringMatching(/^pivot-ui-[\w-]+\.js$/), expect.stringMatching(/^ru-RU-[\w-]+\.js$/), expect.stringMatching(/^typed-entry-[\w-]+\.js$/)])
    expect(new Set(dynamicImports(js))).toEqual(new Set([`./${chunk('pivot-ui')}`, `./${chunk('ru-RU')}`, `./${chunk('typed-entry')}`]))
  })
  it('includes no engine: xlcalc/a1 and xlcalc/text are a few dozen kilobytes', () => {
    expect(js.length).toBeGreaterThan(10_000)
    // the component with its dialogs, bars, panels and PivotTable Fields pane is about 865 KB unminified; the engine would be megabytes
    expect(js.length).toBeLessThan(880_000)
    expect(js).toContain('//#region node_modules/xlcalc/dist/a1.js')
    // xlcalc/text's modules, here the one of stepDecimals (Increase and Decrease Decimal)
    expect(js).toContain('//#region node_modules/xlcalc/dist/numfmt/decimals.js')
  })
  it('exports the public API and, for its chunk, only minified internal names that are not API', async () => {
    const names = [...js.matchAll(/^export \{([^}]*)\}/gm)].flatMap((m) => m[1]!.split(',').map((x) => x.trim().split(/\s+as\s+/).at(-1)!))
    const pub = ['FormatDialog', 'PivotFieldsPane', 'SheetToolbar', 'SheetView', 'TOOLBAR_GROUPS', 'enMessages', 'memorySource', 'ruMessages']
    // the list is the source's: a new or renamed public export is added here on purpose
    expect(Object.keys(await import('../src/index')).sort()).toEqual(pub)
    expect(names.filter((n) => !/^[\w$]{1,2}$/.test(n)).sort()).toEqual(pub)
    // what the PivotTable chunk takes from the main entry point
    const pivot = readFileSync(join(out, chunk('pivot-ui')), 'utf8')
    const taken = [...pivot.matchAll(/import \{([^}]*)\} from "\.\/index\.js"/g)].flatMap((m) => m[1]!.split(',').map((x) => x.trim().split(/\s+as\s+/)[0]!))
    expect(taken.length).toBeGreaterThan(0)
    expect(new Set(names.filter((n) => !pub.includes(n)))).toEqual(new Set(taken))
  })
  it('has no string xlcalc in its code, which a host may look for as a sign of the engine', () => {
    expect(code(js)).not.toMatch(/xlcalc/)
  })
  it('ships with its stylesheet', () => {
    expect(readFileSync(join(out, 'style.css'), 'utf8')).toContain('.xs-sheet')
  })
})

describe('the ru-RU chunk', () => {
  it("contains only ru-RU's formula text, with no import and no string xlcalc in its code", () => {
    const ru = readFileSync(join(out, chunk('ru-RU')), 'utf8')
    expect(ru).toContain('СУММ')
    expect(js).not.toContain('СУММ')
    expect(staticImports(ru)).toEqual([])
    expect(dynamicImports(ru)).toEqual([])
    expect(code(ru)).not.toMatch(/xlcalc/)
    // Since xlcalc 0.42 the locale includes each function's description in Russian, for Insert Function.
    expect(ru.length).toBeLessThan(75_000)
  })
})

describe('the xlcalc/entry chunk', () => {
  it('is bundled separately and in full: no import, so the main entry point loads nothing more for it, and no string xlcalc in its code', () => {
    const entry = readFileSync(join(out, chunk('typed-entry')), 'utf8')
    expect(entry).toMatch(/export \{[^}]*\bentryText\b[^}]*\}/)
    expect(staticImports(entry)).toEqual([])
    expect(dynamicImports(entry)).toEqual([])
    expect(code(entry)).not.toMatch(/xlcalc/)
    // xlcalc's typed-entry module without the engine, including the parts it shares with xlcalc/a1 and xlcalc/text
    expect(entry.length).toBeLessThan(150_000)
  })
})

describe('the PivotTable chunk', () => {
  it('holds the dialogs, menus, buttons, slicers and timelines over PivotTables, and imports only vue and what it shares with the main entry point', () => {
    const pivot = readFileSync(join(out, chunk('pivot-ui')), 'utf8')
    for (const name of ['PivotOptionsDialog', 'PivotFilterMenu', 'SlicerDialogs', 'TimelineBox', 'ToolbarPivot', 'PivotButtons', 'SheetObjects']) {
      expect(pivot).toContain(`__name: "${name}"`)
      expect(js).not.toContain(`__name: "${name}"`)
    }
    // the PivotTable Fields pane is the main entry point's own export
    expect(js).toContain('__name: "PivotFieldsPane"')
    expect(new Set(staticImports(pivot))).toEqual(new Set(['vue', './index.js']))
    expect(dynamicImports(pivot)).toEqual([])
    expect(code(pivot)).not.toMatch(/xlcalc/)
    expect(pivot.length).toBeLessThan(110_000)
  })
})
