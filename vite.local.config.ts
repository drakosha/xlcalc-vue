import { resolve } from 'node:path'
import { defineConfig } from 'vite'

// The `./local` entry, built apart from the main one so that dist/index.js shares no chunk with it.
export default defineConfig({
  build: {
    emptyOutDir: false,
    lib: {
      entry: { local: resolve(import.meta.dirname, 'src/local/index.ts') },
      formats: ['es'],
    },
    rolldownOptions: {
      // the host installs xlcalc: the engine, its worker pool and its xlsx reader stay its own
      external: ['vue', /^xlcalc(\/.*)?$/],
    },
  },
})
