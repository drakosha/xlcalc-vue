// The built `./remote` and `./server` entries need nothing but each other's table of methods: no vue, no engine and no
// string xlcalc in their code, so a page with a source on a server loads no engine and a server needs no Vue.
import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { build } from 'vite'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

const root = resolve(import.meta.dirname, '..')
let out = ''

beforeAll(async () => {
  out = mkdtempSync(join(tmpdir(), 'xs-dist-remote-'))
  const env = process.env.NODE_ENV
  process.env.NODE_ENV = 'production'
  try {
    await build({ configFile: resolve(root, 'vite.remote.config.ts'), root, mode: 'production', logLevel: 'silent', build: { outDir: out, emptyOutDir: true } })
  } finally {
    process.env.NODE_ENV = env
  }
}, 60_000)
afterAll(() => rmSync(out, { recursive: true, force: true }))

const imports = (code: string) => [...code.matchAll(/(?:^|[;\n])\s*import\s*(?:[^'";]*?from\s*)?["']([^"']+)["']|import\(\s*["']([^"']+)["']\s*\)/g)].map((m) => m[1] ?? m[2])
const code = (text: string) => text.replace(/^\/\/#(?:end)?region.*$/gm, '')

describe('dist/remote.js and dist/server.js', () => {
  it('import only their shared chunk, with no string xlcalc in their code', () => {
    const files = readdirSync(out).filter((f) => f.endsWith('.js'))
    expect(files.sort()).toEqual([expect.stringMatching(/^methods-.+\.js$/), 'remote.js', 'server.js'])
    for (const f of files) {
      const text = readFileSync(join(out, f), 'utf8')
      expect(imports(text).every((s) => /^\.\/methods-.+\.js$/.test(s!))).toBe(true)
      expect(code(text)).not.toMatch(/xlcalc|from ["']vue["']/)
      expect(text.length).toBeLessThan(12_000)
    }
  })
})
