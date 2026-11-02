// A workbook's locale whose module fails to load (a chunk the network lost): nothing is read or written as en-US in its
// place, the next request tries again, and ./local fails to open with the import's error rather than open without it.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, SheetToolbar } from '../src'
import { localSource } from '../src/local'
import { loadBookLocale } from '../src/locales'
import { editorKey, goTo, key, mountSheet, settle } from './helpers'

/** How many imports of ru-RU's module fail from now on. */
const lost = vi.hoisted(() => ({ imports: 0 }))
vi.mock('xlcalc/locales/ru-RU', async (real) => {
  if (lost.imports > 0) {
    lost.imports--
    throw new Error('offline')
  }
  return real()
})

afterEach(() => {
  document.body.innerHTML = ''
})

async function open(w: VueWrapper, place: string): Promise<string> {
  await goTo(w, place)
  await key(w, 'F2')
  await settle()
  const text = (w.find('.xs-editor').element as HTMLTextAreaElement).value
  await editorKey(w, 'Escape')
  return text
}

// the tests share the module's state: each before the last keeps ru-RU from loading; vitest gives the import's error as the cause of its own
describe("a ru-RU workbook whose locale's module fails to load", () => {
  it('opens a cell with its displayed text, not the en-US entry text an Enter would send back', async () => {
    lost.imports = 100
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: { value: 1.5, text: '1,5' } } }] }), { locale: 'ru-RU' })
    await settle(12)
    expect(await open(w, 'A1')).toBe('1,5')
  })

  it("leaves the Number Format box without a list, so that en-US's Currency writes no $ code into the book", async () => {
    lost.imports = 100
    const source = { locale: 'ru-RU', sheets: async () => [{ name: 'S', maxRow: 9, maxCol: 9 }], window: async () => ({ pending: { elapsedMs: 0 } }), format: async () => ({ ok: true as const, changes: [] }) }
    const w = mount(SheetToolbar, { props: { source, state: { sheet: 'S', ref: 'A1', cellValue: 1.5 } }, attachTo: document.body })
    await settle(12)
    expect(w.find('[data-tb="number"]').exists()).toBe(false)
    w.unmount()
  })

  it("fails ./local's opening with the import's error", async () => {
    lost.imports = 100
    const s = localSource({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }, { calc: 'inline' })
    await expect(s.ready).rejects.toMatchObject({ cause: { message: 'offline' } })
    await s.close()
  })

  it('tries the module again on the next request', async () => {
    lost.imports = 1
    await expect(loadBookLocale('ru-RU')).rejects.toMatchObject({ cause: { message: 'offline' } })
    expect((await loadBookLocale('ru-RU'))?.tag).toBe('ru-RU')
  })
})
