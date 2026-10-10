// The built main entry carries no engine: the app's guard looks for the string `xlcalc` in its bundle.
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { build } from 'vite'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

const root = resolve(import.meta.dirname, '..')
let out = ''
let js = ''

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
}, 60_000)
afterAll(() => rmSync(out, { recursive: true, force: true }))

describe('dist/index.js', () => {
  it('has no string xlcalc at all', () => {
    expect(js.length).toBeGreaterThan(10_000)
    expect(js.match(/xlcalc/gi)).toBeNull()
  })
  it('imports nothing but vue', () => {
    const imports = [...js.matchAll(/(?:^|[;\n])\s*import\s*(?:[^'";]*?from\s*)?["']([^"']+)["']/g)].map((m) => m[1])
    expect(imports.length).toBeGreaterThan(0)
    expect(new Set(imports)).toEqual(new Set(['vue']))
    expect(js).not.toMatch(/import\(/)
  })
  it('comes with its style sheet', () => {
    expect(readFileSync(join(out, 'style.css'), 'utf8')).toContain('.xs-sheet')
  })
})
