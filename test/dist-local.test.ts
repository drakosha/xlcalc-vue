// The built `./local` entry point: its own file next to the main one, which takes nothing from it; xlcalc stays the host's.
import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { build } from 'vite'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

const root = resolve(import.meta.dirname, '..')
let out = ''

beforeAll(async () => {
  out = mkdtempSync(join(tmpdir(), 'xs-dist-local-'))
  const env = process.env.NODE_ENV
  process.env.NODE_ENV = 'production'
  try {
    // as `npm run build` produces them: the main entry point first, then the local one next to it
    await build({ configFile: resolve(root, 'vite.config.ts'), root, mode: 'production', logLevel: 'silent', build: { outDir: out, emptyOutDir: true } })
    await build({ configFile: resolve(root, 'vite.local.config.ts'), root, mode: 'production', logLevel: 'silent', build: { outDir: out } })
  } finally {
    process.env.NODE_ENV = env
  }
}, 60_000)
afterAll(() => rmSync(out, { recursive: true, force: true }))

const importsOf = (js: string) => [...js.matchAll(/(?:^|[;\n])\s*import\s*(?:[^'";]*?from\s*)?["']([^"']+)["']|import\(\s*["']([^"']+)["']\s*\)/g)].map((m) => m[1] ?? m[2])

describe('dist/local.js', () => {
  it('is one file next to index.js and its ru-RU and xlcalc/entry chunks, which import no xlcalc', () => {
    const files = readdirSync(out).filter((f) => f.endsWith('.js'))
    const ru = files.find((f) => /^ru-RU-[\w-]+\.js$/.test(f))
    const entry = files.find((f) => /^typed-entry-[\w-]+\.js$/.test(f))
    expect(files.sort()).toEqual(['index.js', 'local.js', ru, entry].sort())
    const index = readFileSync(join(out, 'index.js'), 'utf8')
    expect(new Set(importsOf(index))).toEqual(new Set(['vue', `./${ru}`, `./${entry}`]))
  })
  it('imports xlcalc and loads the pool, the xlsx reader, a locale and the function texts only when requested', () => {
    const local = readFileSync(join(out, 'local.js'), 'utf8')
    expect(new Set(importsOf(local))).toEqual(new Set(['xlcalc', 'xlcalc/a1', 'xlcalc/text', 'xlcalc/worker', 'xlcalc/xlsx', 'xlcalc/locales/ru-RU', 'xlcalc/function-text', 'xlcalc/function-text/ru-RU']))
    expect(local).toMatch(/import\(\s*["']xlcalc\/worker["']\s*\)/)
    expect(local).toMatch(/import\(\s*["']xlcalc\/locales\/ru-RU["']\s*\)/)
    expect(local).toMatch(/import\(\s*["']xlcalc\/function-text["']\s*\)/)
    expect(local).not.toMatch(/from\s*["']xlcalc\/(worker|xlsx|locales\/[^"']+|function-text[^"']*)["']/)
  })
})
