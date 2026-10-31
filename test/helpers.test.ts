// The test helpers themselves: goTo right after mounting waits for the sheets, which a cold source reads after mountSheet settles.
import { afterEach, expect, it } from 'vitest'
import { localSource, type LocalSource } from '../src/local'
import { goTo, mountSheet, until } from './helpers'
import { slicerBook } from './slicer-fixture'

let source: LocalSource | null = null
afterEach(async () => {
  document.body.innerHTML = ''
  await source?.close()
})

it('goTo to another sheet right after mounting goes there when the sheets are read late', async () => {
  source = localSource(slicerBook(), { calc: 'inline' })
  const sheets = source.sheets.bind(source)
  // the sheets arrive well after mountSheet's rounds, as on a busy machine the first test of a file loads its modules
  source.sheets = async (signal) => {
    await new Promise((r) => setTimeout(r, 200))
    return sheets(signal)
  }
  const w = await mountSheet(source)
  await goTo(w, 'Report!B4')
  await until(() => !!document.querySelector('[data-xs-object]'), 3000)
  expect(document.querySelector('[role="tab"][aria-selected="true"]')!.textContent!.trim()).toBe('Report')
})
