import { resolve } from 'node:path'
import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'

// The example app (`npm run dev:example`); built into example/dist, which git ignores.
export default defineConfig({
  root: resolve(import.meta.dirname, 'example'),
  plugins: [vue()],
  server: { port: 5181 },
  build: { outDir: resolve(import.meta.dirname, 'example/dist'), emptyOutDir: true },
})
