import { copyFileSync } from 'node:fs'
import { resolve } from 'node:path'
import vue from '@vitejs/plugin-vue'
import { defineConfig, type Plugin } from 'vite'

// The example app (`npm run dev:example`): the "all formatting" book (index.html), the local source (local.html) and the
// app like Excel (spa.html); built into example/dist, which git ignores.
// `--mode demo` (`npm run build:demo`) builds the app alone for xlcalc.pridelab.dev/demo/ into example/dist-demo: its
// paths under /demo/, its index.html the app itself, no source maps.

const example = (file: string) => resolve(import.meta.dirname, 'example', file)

/** The demo's index.html: a copy of the built spa.html, so that /demo/ opens the app. */
const demoIndex = (): Plugin => ({
  name: 'xs-demo-index',
  apply: 'build',
  writeBundle(opts) {
    copyFileSync(resolve(opts.dir!, 'spa.html'), resolve(opts.dir!, 'index.html'))
  },
})

export default defineConfig(({ mode }) => {
  const demo = mode === 'demo'
  const input: Record<string, string> = demo ? { spa: example('spa.html') } : { index: example('index.html'), local: example('local.html'), spa: example('spa.html') }
  return {
    root: example(''),
    base: demo ? '/demo/' : '/',
    plugins: [vue(), ...(demo ? [demoIndex()] : [])],
    server: { port: 5181 },
    build: {
      outDir: example(demo ? 'dist-demo' : 'dist'),
      emptyOutDir: true,
      sourcemap: false,
      rolldownOptions: { input },
    },
  }
})
