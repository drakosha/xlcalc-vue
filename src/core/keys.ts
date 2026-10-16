// Shortcut keys on any layout and system: the key's Latin letter, a Mac, Control or Cmd held.

/** A shortcut's Latin letter: the key's own, else its place on the keyboard (`KeyB` under Russian `и`); '' for anything else. */
export function shortcutLetter(e: KeyboardEvent): string {
  if (/^[a-z]$/i.test(e.key)) return e.key.toLowerCase()
  return /^Key[A-Z]$/.test(e.code) ? e.code.slice(3).toLowerCase() : ''
}

/** A Mac, where Excel ends a copy's text without a line break (Windows with CRLF), drags a series with Option and names Cmd in tips. */
export const onMac = (): boolean => typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform)

/** Control or Cmd held: Excel's shortcuts take either, Cmd on a Mac. */
export const isMod = (e: KeyboardEvent | MouseEvent): boolean => e.ctrlKey || e.metaKey

/** A shortcut and not a character: Control alone or Cmd; AltGr comes as Control+Alt and types (`@`, `[` on many layouts). */
export const isShortcut = (e: KeyboardEvent): boolean => (e.ctrlKey && !e.altKey) || e.metaKey
