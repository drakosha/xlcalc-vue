import { createHash } from 'node:crypto'
import { createRequire } from 'node:module'
import { resolve } from 'node:path'
import vue from '@vitejs/plugin-vue'
import { build, type Plugin, type Rollup } from 'vite'
import { defineConfig } from 'vitest/config'

/**
 * xlcalc/entry as a chunk of its own, bundled apart: the modules it shares with xlcalc/a1 and xlcalc/text stay in it, so
 * the main entry loads nothing more for it, and a source with `editText` never loads it. Bundled with the main entry it
 * would pull its part of those shared modules into a chunk the main entry imports at once.
 */
function standaloneEntry(): Plugin {
  let file = ''
  let code = ''
  return {
    name: 'xs-standalone-entry',
    apply: 'build',
    enforce: 'pre',
    async buildStart() {
      const entry = createRequire(import.meta.url).resolve('xlcalc/entry')
      const out = (await build({
        configFile: false,
        logLevel: 'silent',
        mode: 'production',
        build: { write: false, lib: { entry, formats: ['es'], fileName: 'entry' } },
      })) as Rollup.RollupOutput[]
      code = out[0]!.output[0].code
      file = `typed-entry-${createHash('sha256').update(code).digest('base64url').slice(0, 8)}.js`
    },
    resolveId: (id) => (id === 'xlcalc/entry' ? { id: `./${file}`, external: true } : null),
    generateBundle() {
      this.emitFile({ type: 'asset', fileName: file, source: code })
    },
  }
}

export default defineConfig({
  plugins: [vue(), standaloneEntry()],
  build: {
    lib: {
      entry: { index: resolve(import.meta.dirname, 'src/index.ts') },
      formats: ['es'],
      cssFileName: 'style',
    },
    rolldownOptions: {
      // vue is the host's, xlcalc and its contract are types only; xlcalc/a1, xlcalc/text and a locale's chunk are bundled
      external: ['vue', 'xlcalc', 'xlcalc/contract'],
    },
  },
  test: {
    environment: 'happy-dom',
    include: ['test/**/*.test.ts'],
    setupFiles: ['test/setup.ts'],
    passWithNoTests: true,
    // A test takes well under a second alone; the limit guards against a hang, not a slow machine
    testTimeout: 20_000,
  },
})
