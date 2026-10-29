// @vitest-environment node
// Excel for Windows' order of a dialog's buttons: the main one (OK, Insert, Delete) before Cancel, in every dialog's markup.
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const DIR = join(import.meta.dirname, '../src/components')
const isButton = (line: string) => line.includes('<button')
const isCancel = (line: string) => isButton(line) && /\{\{ (inserted \? m\.close : )?m\.cancel \}\}/.test(line)
const isMain = (line: string) => isButton(line) && line.includes('xs-button-main')

describe("dialogs' buttons", () => {
  it('put the main button before Cancel, as Excel for Windows does', () => {
    const wrong: string[] = []
    for (const file of readdirSync(DIR).filter((f) => f.endsWith('.vue'))) {
      const lines = readFileSync(join(DIR, file), 'utf8').split('\n')
      lines.forEach((line, i) => {
        if (!isCancel(line)) return
        // a main button after Cancel in the same row of buttons
        for (let j = i + 1; j < lines.length && isButton(lines[j]!); j++) if (isMain(lines[j]!)) wrong.push(`${file}:${i + 1}`)
      })
    }
    expect(wrong).toEqual([])
  })
})
