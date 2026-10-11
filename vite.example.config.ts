import { resolve } from 'node:path'
import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'

// The example app (`npm run dev:example`): the "all formatting" book (index.html), the local source (local.html) and the
// app like Excel (spa.html); built into example/dist, which git ignores.
export default defineConfig({
  root: resolve(import.meta.dirname, 'example'),
  plugins: [vue()],
  server: { port: 5181 },
  build: {
    outDir: resolve(import.meta.dirname, 'example/dist'),
    emptyOutDir: true,
    rolldownOptions: { input: { index: resolve(import.meta.dirname, 'example/index.html'), local: resolve(import.meta.dirname, 'example/local.html'), spa: resolve(import.meta.dirname, 'example/spa.html') } },
  },
})
