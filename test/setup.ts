// Node 26 has its own global `localStorage` (empty without a file) that hides happy-dom's: give the tests a working one.
class MemoryStorage {
  private items = new Map<string, string>()
  get length(): number {
    return this.items.size
  }
  key(i: number): string | null {
    return [...this.items.keys()][i] ?? null
  }
  getItem(key: string): string | null {
    return this.items.get(key) ?? null
  }
  setItem(key: string, value: string): void {
    this.items.set(key, String(value))
  }
  removeItem(key: string): void {
    this.items.delete(key)
  }
  clear(): void {
    this.items.clear()
  }
}

if (typeof globalThis.localStorage?.getItem !== 'function') {
  Object.defineProperty(globalThis, 'localStorage', { value: new MemoryStorage(), configurable: true, writable: true })
}
