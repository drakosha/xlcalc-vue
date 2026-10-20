// After vue-tsc: each relative import in dist's declarations names its file with `.js`, as Node's ESM resolution requires
// (`./contract` → `./contract.js`, `./SheetView.vue` → `./SheetView.vue.js`, `..` → `../index.js`), and stylesheet imports
// are dropped, since a declaration has no styles. Run by `npm run build`.
import { existsSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { dirname, join, relative, resolve } from 'node:path'

const dist = resolve(import.meta.dirname, '..', 'dist')
const files = readdirSync(dist, { recursive: true, encoding: 'utf8' }).filter((f) => f.endsWith('.d.ts')).map((f) => join(dist, f))

/** The `.js` specifier of the declaration a relative specifier resolves to, from the directory `from`. */
function target(from: string, spec: string): string {
  const path = resolve(from, spec.replace(/\.[jt]s$/, ''))
  const file = [`${path}.d.ts`, join(path, 'index.d.ts')].find(existsSync)
  if (!file) throw new Error(`dts-paths: ${spec} in ${relative(dist, from)} names no declaration`)
  const out = relative(from, file).replace(/\.d\.ts$/, '.js')
  return out.startsWith('.') ? out : `./${out}`
}

let changed = 0
for (const file of files) {
  const from = dirname(file)
  const text = readFileSync(file, 'utf8')
  const next = text
    .replace(/^import ['"][^'"]+\.css['"];?\n/gm, '')
    .replace(/(\bfrom |\bimport\()(['"])(\.\.?(?:\/[^'"]*)?)\2/g, (_, head: string, q: string, spec: string) => `${head}${q}${target(from, spec)}${q}`)
  if (next !== text) {
    writeFileSync(file, next)
    changed++
  }
}
console.log(`dts-paths: ${changed} of ${files.length} declarations rewritten`)
