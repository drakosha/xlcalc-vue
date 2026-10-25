// Storage for the frozen panes display choice, Show Formulas and column widths: the browser's storage under the host's key, otherwise for the component's lifetime.

export interface Memory {
  recall<T>(name: string): T | null
  remember(name: string, value: unknown): void
}

/** Storage over `localStorage` under `storageKey()`, or an in-memory map when the host provides no key. */
export function createMemory(storageKey: () => string | undefined): Memory {
  const own = new Map<string, string>()
  return {
    recall<T>(name: string): T | null {
      try {
        const key = storageKey()
        const raw = key ? localStorage.getItem(`${key}:${name}`) : (own.get(name) ?? null)
        return raw === null ? null : (JSON.parse(raw) as T)
      } catch {
        return null
      }
    },
    remember(name: string, value: unknown): void {
      const raw = value === null ? null : JSON.stringify(value)
      const key = storageKey()
      if (!key) {
        if (raw === null) own.delete(name)
        else own.set(name, raw)
        return
      }
      try {
        if (raw === null) localStorage.removeItem(`${key}:${name}`)
        else localStorage.setItem(`${key}:${name}`, raw)
      } catch {
        // Private mode or denied storage: the choice lasts until the page is closed.
      }
    },
  }
}
