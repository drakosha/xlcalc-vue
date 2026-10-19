// The API reference keeps up with the code: every export, source method, sheet method, prop, event and CSS variable a host
// can use is named in docs/api.md, and the relative links of the README and docs/ lead to files that exist.
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const root = resolve(import.meta.dirname, '..')
const read = (path: string) => readFileSync(join(root, path), 'utf8')
const api = read('docs/api.md')

/** Names given in backticks in the reference: `Name`, `name(...)`, `name?`. */
const named = new Set([...api.matchAll(/`([\w-]+)/g)].map((m) => m[1]!))
const missing = (names: Iterable<string>) => [...new Set(names)].filter((n) => !named.has(n)).sort()

/** The body of `export interface Name { ... }` in a file, up to its closing brace at the line's start. */
function body(path: string, start: RegExp): string {
  const text = read(path)
  const at = text.search(start)
  if (at < 0) throw new Error(`${start} not found in ${path}`)
  const from = text.indexOf('{', at) + 1
  return text.slice(from, text.indexOf('\n}', from))
}

/** Members of an interface or an inline props/emits type: `name(`, `name?:`, `readonly name:`, `name: [` at the block's own indent. */
function members(block: string): string[] {
  const indent = Math.min(...[...block.matchAll(/^( +)\w/gm)].map((m) => m[1]!.length))
  const re = new RegExp(`^ {${indent}}(?:readonly )?(\\w+)\\??[(:]`, 'gm')
  return [...block.matchAll(re)].map((m) => m[1]!)
}

/** The names an entry exports: `export { a, b as c, type D }` and `export type { ... }`. */
function exportsOf(path: string): string[] {
  const names: string[] = []
  for (const m of read(path).matchAll(/export (?:type )?\{([^}]*)\}/g)) {
    for (const part of m[1]!.split(',')) {
      const name = part.trim().replace(/^type /, '').split(/ as /).pop()?.trim()
      if (name) names.push(name)
    }
  }
  return names
}

/** Component-private variables: what the sheet sets itself or keeps fixed, which a host does not override. */
const INTERNAL_VARS = new Set(['--xs-fx-lines', '--xs-explain-depth', '--xs-outline-h', '--xs-outline-w', '--xs-swatch', '--xs-touch-target'])

describe('docs/api.md', () => {
  it('names every export of both entries and every type of the contract', () => {
    const contract = readdirSync(join(root, 'src/contract')).flatMap((f) => [...read(`src/contract/${f}`).matchAll(/^export (?:interface|type) (\w+)/gm)].map((m) => m[1]!))
    expect(missing([...exportsOf('src/index.ts'), ...exportsOf('src/local/index.ts'), ...contract])).toEqual([])
  })

  it('names every method of the source, the sheet and the local source', () => {
    const names = [
      ...members(body('src/contract/source.ts', /export interface SheetSource\b/)),
      ...members(body('src/contract/view.ts', /export interface SheetViewApi\b/)),
      ...members(body('src/local/types.ts', /export interface LocalSource\b/)),
      ...members(body('src/local/types.ts', /export interface LocalSourceOptions\b/)),
      ...members(body('src/contract/memory.ts', /export interface MemorySource\b/)),
    ]
    expect(missing(names)).toEqual([])
  })

  it('names every prop and event of the components', () => {
    const names = members(body('src/contract/view.ts', /export interface SheetViewProps\b/))
    for (const file of ['SheetView', 'SheetToolbar', 'FormatDialog']) {
      const vue = `src/components/${file}.vue`
      if (file !== 'SheetView') names.push(...members(body(vue, /defineProps<\{/)))
      names.push(...members(body(vue, /defineEmits<\{/)))
    }
    expect(missing(names)).toEqual([])
  })

  it('names every CSS variable a host can set', () => {
    const css = readdirSync(join(root, 'src'))
      .filter((f) => f.endsWith('.css'))
      .map((f) => read(`src/${f}`))
      .join('\n')
    const vars = [...css.matchAll(/(--xs-[\w-]+):|var\((--xs-[\w-]+),/g)].map((m) => (m[1] ?? m[2])!).filter((v) => !INTERNAL_VARS.has(v))
    expect(missing(vars)).toEqual([])
  })
})

describe('links of the README and docs/', () => {
  it('lead to files that exist', () => {
    const pages = ['README.md', ...readdirSync(join(root, 'docs')).filter((f) => f.endsWith('.md')).map((f) => `docs/${f}`)]
    const broken = pages.flatMap((page) =>
      [...read(page).matchAll(/\]\(([^)#\s]+)(?:#[^)]*)?\)/g)]
        .map((m) => m[1]!)
        .filter((target) => !/^[a-z]+:/.test(target) && !existsSync(resolve(root, dirname(page), target)))
        .map((target) => `${page} → ${target}`),
    )
    expect(broken).toEqual([])
  })
})
