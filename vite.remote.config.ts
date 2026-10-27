import { resolve } from 'node:path'
import { defineConfig } from 'vite'

// The `./remote` and `./server` entries, built apart from the main one: neither needs vue or the engine, only each other's table of methods.
export default defineConfig({
  build: {
    emptyOutDir: false,
    lib: {
      entry: { remote: resolve(import.meta.dirname, 'src/remote/index.ts'), server: resolve(import.meta.dirname, 'src/remote/server.ts') },
      formats: ['es'],
    },
    rolldownOptions: {
      external: ['vue', /^xlcalc(\/.*)?$/],
    },
  },
})
