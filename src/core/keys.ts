// Shortcut keys on any keyboard layout and system: the key's Latin letter, whether it is a Mac, Ctrl or Cmd held.

/** A shortcut's Latin letter: the key's own, otherwise its position on the keyboard (`KeyB` under Russian `и`); '' for anything else. */
export function shortcutLetter(e: KeyboardEvent): string {
  if (/^[a-z]$/i.test(e.key)) return e.key.toLowerCase()
  return /^Key[A-Z]$/.test(e.code) ? e.code.slice(3).toLowerCase() : ''
}

/** Whether it is a Mac, where Excel ends a copy's text without a line break (Windows with CRLF), drags a series with Option and names Cmd in tooltips. */
export const onMac = (): boolean => typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform)

/** Ctrl or Cmd held: Excel's shortcuts accept either, Cmd on a Mac. */
export const isMod = (e: KeyboardEvent | MouseEvent): boolean => e.ctrlKey || e.metaKey

/** A shortcut and not a character: Ctrl alone or Cmd; AltGr arrives as Ctrl+Alt and types characters (`@`, `[` on many layouts). */
export const isShortcut = (e: KeyboardEvent): boolean => (e.ctrlKey && !e.altKey) || e.metaKey
