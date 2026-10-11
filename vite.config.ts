import { resolve } from 'node:path'
import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  plugins: [vue()],
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
  },
})
