// The example app's notice of what a file lost, as Excel's message bar: one line with each kind and its count, the
// details (messages, sheets, cells, names) on a click, and a close button.
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { XlsxIssue } from 'xlcalc/xlsx'
import SpaLosses from '../example/SpaLosses.vue'
import { lossesOf } from '../example/spa-open'
import { TEXTS } from '../example/spa-texts'

afterEach(() => {
  document.body.innerHTML = ''
})

const t = TEXTS.ru
const hidden = (n: number): XlsxIssue[] => Array.from({ length: n }, (_, i) => ({ code: 'name', message: 'скрытое имя прочитано как видимое', name: `n${i}` }))
const issues: XlsxIssue[] = [
  ...hidden(28_396),
  { code: 'externalLink', message: 'связи с другими книгами не поддерживаются' },
  { code: 'externalLink', message: 'ссылки на другие книги', sheet: 'Лист 1', cells: ['A1', 'A2', 'A3'] },
  { code: 'externalLink', message: 'ссылки на другие книги', sheet: 'Data', cells: ['E5', 'E6', 'E7', 'E8'] },
  ...Array.from({ length: 10 }, (_, i): XlsxIssue => ({ code: 'drawing', message: 'рисунки не поддерживаются', sheet: `S${i + 1}` })),
  ...Array.from({ length: 7 }, (_, i): XlsxIssue => ({ code: 'style', message: `стиль ${i + 1}`, sheet: 'Data', cells: [`B${i + 1}`] })),
]
const mountNotice = () => mount(SpaLosses, { props: { kinds: lossesOf(issues), t }, attachTo: document.body })

describe('SpaLosses: what a file lost, in one line', () => {
  it('names each kind with its count in one line, what changes values first, without the details', () => {
    const w = mountNotice()
    expect(w.find('[data-spa-lost-title]').text()).toBe('Не перенесено')
    expect(w.find('[data-spa-lost-summary]').text()).toBe(`связи с другими книгами (3), рисунки и диаграммы (10), оформление (7), имена (${t.count(28_396)})`)
    expect(t.count(28_396)).toBe('28 396')
    expect(w.find('[data-spa-lost-details]').exists()).toBe(false)
    expect(w.find('[data-spa-lost-more]').attributes('aria-expanded')).toBe('false')
  })

  it('shows the details on a click: messages with their counts, sheets, cells and names, the long lists cut', async () => {
    const w = mountNotice()
    await w.find('[data-spa-lost-more]').trigger('click')
    const details = w.find('[data-spa-lost-details]')
    expect(details.exists()).toBe(true)
    expect(w.find('[data-spa-lost-more]').attributes('aria-expanded')).toBe('true')
    expect(w.find('[data-spa-lost-more]').text()).toBe('Скрыть подробности')
    const kinds = details.findAll('[data-spa-lost-kind]')
    expect(kinds.map((k) => k.attributes('data-spa-lost-kind'))).toEqual(['externalLink', 'drawing', 'style', 'name'])
    const link = kinds[0]!
    expect(link.find('h3').text()).toBe('связи с другими книгами (3)')
    expect(link.findAll('li').map((li) => li.text())).toEqual(['связи с другими книгами не поддерживаются', 'ссылки на другие книги (2)'])
    expect(link.text()).toContain('Листы: Лист 1, Data')
    expect(link.text()).toContain("Ячейки: 'Лист 1'!A1, 'Лист 1'!A2, 'Лист 1'!A3, Data!E5, Data!E6, всего 7")
    expect(kinds[1]!.text()).toContain('Листы: S1, S2, S3, S4, S5, S6, S7, S8 и ещё 2')
    expect(kinds[2]!.findAll('li').map((li) => li.text())).toEqual(['стиль 1', 'стиль 2', 'стиль 3', 'стиль 4', 'стиль 5', 'и ещё 2'])
    expect(kinds[3]!.text()).toContain(`Имена: n0, n1, n2, n3, n4, всего ${t.count(28_396)}`)
    await w.find('[data-spa-lost-more]').trigger('click')
    expect(w.find('[data-spa-lost-details]').exists()).toBe(false)
  })

  it('closes from its button', async () => {
    const w = mountNotice()
    await w.find('[data-spa-lost-close]').trigger('click')
    expect(w.emitted('close')).toHaveLength(1)
  })
})
