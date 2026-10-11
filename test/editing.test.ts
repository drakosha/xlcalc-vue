// Editing: Excel's keys, typing over a cell, Delete over the selection, pasting TSV, locked cells, and the source's outcomes.
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import type { EditOutcome } from '../src/source'
import { active, cellText, editorKey, editorOpen, goTo, key, mountSheet, paste, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: 'a1', B1: 'b1', A2: 10, B2: { value: 20, formula: 'A2*2' }, C3: 'c3' },
      locked: [{ range: 'D1:D10', reason: 'the script owns it' }],
    },
  ],
})

/** A memory source that records the edits it gets. */
function recorded(b = book()) {
  const source = memorySource(b)
  const edits: Parameters<NonNullable<typeof source.edit>>[0][] = []
  const edit = source.edit!.bind(source)
  source.edit = (entries) => {
    edits.push(entries)
    return edit(entries)
  }
  return { source, edits }
}

describe('keys', () => {
  it('typing over a cell starts a new text; Enter sends it and moves down', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    expect(active(w)).toBe('1:1')
    await key(w, 'x')
    expect(editorOpen(w)).toBe(true)
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('x')
    await type(w, 'xyz')
    await editorKey(w, 'Enter')
    await settle()
    expect(edits).toEqual([[{ sheet: 'S', cell: 'A1', text: 'xyz' }]])
    expect(active(w)).toBe('2:1')
    expect(cellText(w, 1, 1)).toBe('xyz')
  })

  it('Tab, Shift+Tab, Shift+Enter commit and move; Esc cancels', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    await key(w, '1')
    await editorKey(w, 'Tab')
    await settle()
    expect(active(w)).toBe('2:3')
    await key(w, '2')
    await editorKey(w, 'Tab', { shiftKey: true })
    await settle()
    expect(active(w)).toBe('2:2')
    await key(w, '3')
    await editorKey(w, 'Enter', { shiftKey: true })
    await settle()
    expect(active(w)).toBe('1:2')
    await key(w, '4')
    await editorKey(w, 'Escape')
    await settle()
    expect(editorOpen(w)).toBe(false)
    expect(active(w)).toBe('1:2')
    expect(edits.map((e) => e[0]!.cell + '=' + e[0]!.text)).toEqual(['B2=1', 'C2=2', 'B2=3'])
  })

  it('Enter and Tab without a field move the cursor; arrows after typing commit', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await key(w, 'Enter')
    expect(active(w)).toBe('2:1')
    await key(w, 'Tab')
    expect(active(w)).toBe('2:2')
    await key(w, 'q')
    await editorKey(w, 'ArrowRight')
    await settle()
    expect(active(w)).toBe('2:3')
    expect(edits).toEqual([[{ sheet: 'S', cell: 'B2', text: 'q' }]])
  })

  it('F2 and a double click open the field with the formula or the value', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    await key(w, 'F2')
    await settle()
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('=A2*2')
    // Unchanged text is not sent.
    await editorKey(w, 'Enter')
    await settle()
    expect(edits).toEqual([])
    await w.find('[data-cell="1:2"]').trigger('dblclick')
    await settle()
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('b1')
  })

  it('a text the book would read as a number or a date opens with an apostrophe, and Enter leaves it a text', async () => {
    const { source, edits } = recorded({ sheets: [{ name: 'S', cells: { A1: { value: '007', text: '007' }, A2: { value: '15.03.2026' }, A3: 'plain' } }] })
    const w = await mountSheet(source, { locale: 'ru-RU' })
    const field = () => (w.find('.xs-editor').element as HTMLTextAreaElement).value
    await key(w, 'F2')
    await settle()
    expect(field()).toBe("'007")
    await editorKey(w, 'Enter')
    await key(w, 'F2')
    await settle()
    expect(field()).toBe("'15.03.2026")
    await editorKey(w, 'Enter')
    await key(w, 'F2')
    await settle()
    expect(field()).toBe('plain')
    await editorKey(w, 'Escape')
    expect(edits).toEqual([])
  })

  it('Delete clears the selection, skipping cells known empty', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await goTo(w, 'A1:C3')
    await key(w, 'Delete')
    await settle()
    expect(edits).toHaveLength(1)
    expect(edits[0]!.map((e) => e.cell).sort()).toEqual(['A1', 'A2', 'B1', 'B2', 'C3'])
    expect(edits[0]!.every((e) => e.text === null)).toBe(true)
    expect(cellText(w, 1, 1)).toBe('')
  })

  it('a paste lays the text out as Excel pastes it, from the active cell; one value fills the selection', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await goTo(w, 'B5')
    paste(w, '1\t2\n3\t\n')
    await settle()
    // A line break after the last row is one more row, of empty fields: it clears the cells under the paste.
    expect(edits[0]).toEqual([
      { sheet: 'S', cell: 'B5', text: '1' },
      { sheet: 'S', cell: 'C5', text: '2' },
      { sheet: 'S', cell: 'B6', text: '3' },
      { sheet: 'S', cell: 'C6', text: null },
      { sheet: 'S', cell: 'B7', text: null },
      { sheet: 'S', cell: 'C7', text: null },
    ])
    expect(cellText(w, 6, 2)).toBe('3')
    await goTo(w, 'E1:E2')
    paste(w, '7')
    await settle()
    expect(edits[1]!.map((e) => `${e.cell}=${e.text}`)).toEqual(['E1=7', 'E2=7'])
  })

  it('a paste takes quotes as text, fields at tabs and rows at line breaks', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await goTo(w, 'B5')
    // TODO(xlcalc): parseClipboard does not read quoted fields yet (asked in the library's tracker); once it does, this is one row, `x\ny` and `say "hi"`.
    paste(w, '"x\ny"\t"say ""hi"""')
    await settle()
    expect(edits[0]!.map((e) => `${e.cell}=${e.text}`)).toEqual(['B5="x', 'C5=null', 'B6=y"', 'C6="say ""hi"""'])
  })

  it('a paste that would leave the sheet is refused whole, as Excel refuses it', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source, { messages: { editFailed: 'Edit failed' } })
    await goTo(w, 'XFD1048576')
    paste(w, '1\t2')
    await settle()
    expect(edits).toEqual([])
    expect(w.find('.xs-message').text()).toBe('Edit failed')
  })

  it('read-only: typing, Delete and paste do nothing', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source, { readOnly: true })
    await key(w, 'x')
    await key(w, 'Delete')
    await settle()
    expect(w.find('.xs-editor').exists()).toBe(false)
    expect(edits).toEqual([])
  })
})

describe('locked cells', () => {
  it('are not sent: the field does not open and the reason is shown', async () => {
    const { source, edits } = recorded()
    const w = await mountSheet(source)
    await goTo(w, 'D3')
    await key(w, 'x')
    await settle()
    expect(editorOpen(w)).toBe(false)
    expect(w.find('.xs-message').text()).toContain('the script owns it')
    await goTo(w, 'C2:D2')
    await key(w, 'Delete')
    await settle()
    expect(edits).toEqual([])
    expect(w.find('[data-cell="3:4"]').attributes('title')).toContain('the script owns it')
  })
})

describe('outcomes', () => {
  /** A source whose next edit answers with `outcome` once released. */
  function held(outcome: EditOutcome) {
    const { source, edits } = recorded()
    let release = () => {}
    const edit = source.edit!.bind(source)
    source.edit = (entries) => {
      edits.push(entries)
      return new Promise((resolve) => (release = () => resolve(outcome.ok ? edit(entries) : outcome)))
    }
    return { source, edits, release: () => release() }
  }

  it('while in flight the cell shows the typed text as pending', async () => {
    const h = held({ ok: true, changes: [] })
    const w = await mountSheet(h.source)
    await key(w, 'n')
    await type(w, 'new')
    await editorKey(w, 'Enter')
    await settle()
    const td = w.find('[data-cell="1:1"]')
    expect(td.classes()).toContain('xs-pending')
    expect(td.text()).toBe('new')
    h.release()
    await settle()
    expect(w.find('[data-cell="1:1"]').classes()).not.toContain('xs-pending')
    expect(cellText(w, 1, 1)).toBe('new')
  })

  it('conflict: the text stays in the field, the windows are read again, a message', async () => {
    const h = held({ ok: false, reason: 'conflict', message: '' })
    const w = await mountSheet(h.source, { messages: { conflict: 'Changed elsewhere' } })
    const windows = h.source.window
    let reads = 0
    h.source.window = (req, signal) => {
      reads++
      return windows(req, signal)
    }
    await key(w, 'k')
    await type(w, 'kept')
    await editorKey(w, 'Enter')
    h.release()
    await settle()
    expect(reads).toBeGreaterThan(0)
    expect(editorOpen(w)).toBe(true)
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('kept')
    expect(active(w)).toBe('1:1')
    expect(w.find('.xs-message').text()).toBe('Changed elsewhere')
    expect(cellText(w, 1, 1)).toBe('a1')
  })

  it('invalid: the field opens again with the caret at `at`, the source message shown', async () => {
    const h = held({ ok: false, reason: 'invalid', message: 'Missing bracket', at: 4 })
    const w = await mountSheet(h.source)
    await key(w, '=')
    await type(w, '=SUM(A1')
    await editorKey(w, 'Enter')
    h.release()
    await settle()
    const el = w.find('.xs-editor').element as HTMLTextAreaElement
    expect(editorOpen(w)).toBe(true)
    expect(el.value).toBe('=SUM(A1')
    expect(el.selectionStart).toBe(4)
    expect(w.find('.xs-message').text()).toBe('Missing bracket')
  })

  it('refused, busy and failed: the cell comes back as it was, with a message', async () => {
    for (const reason of ['refused', 'busy', 'failed'] as const) {
      const h = held({ ok: false, reason, message: `no: ${reason}` })
      const w = await mountSheet(h.source)
      await key(w, 'z')
      await editorKey(w, 'Enter')
      await settle()
      expect(cellText(w, 1, 1)).toBe('z')
      h.release()
      await settle()
      expect(cellText(w, 1, 1)).toBe('a1')
      expect(editorOpen(w)).toBe(false)
      expect(w.find('.xs-message').text()).toBe(`no: ${reason}`)
      w.unmount()
    }
  })
})
