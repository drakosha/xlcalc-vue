// Reading a zip's entries in tests, without a dependency: the central directory, stored or deflated entries.
import { inflateRawSync } from 'node:zlib'

/** A zip's entries as text, stored or deflated, by the central directory. */
export function unzipText(data: Uint8Array): Record<string, string> {
  const b = Buffer.from(data)
  const end = b.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
  const out: Record<string, string> = {}
  let at = b.readUInt32LE(end + 16)
  for (let i = 0; i < b.readUInt16LE(end + 10); i++) {
    const [method, size, nameLen, extra, comment, local] = [b.readUInt16LE(at + 10), b.readUInt32LE(at + 20), b.readUInt16LE(at + 28), b.readUInt16LE(at + 30), b.readUInt16LE(at + 32), b.readUInt32LE(at + 42)]
    const name = b.toString('utf8', at + 46, at + 46 + nameLen)
    const from = local + 30 + b.readUInt16LE(local + 26) + b.readUInt16LE(local + 28)
    const raw = b.subarray(from, from + size)
    out[name] = (method === 8 ? inflateRawSync(raw) : raw).toString('utf8')
    at += 46 + nameLen + extra + comment
  }
  return out
}
