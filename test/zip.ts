// Reading a zip's entries in tests, and writing a stored one, without a dependency.
import { crc32, inflateRawSync } from 'node:zlib'

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

/** A zip of text entries, stored, for a file a test edits by hand (an xlsx with what xlcalc does not write). */
export function zipText(files: Record<string, string>): Uint8Array {
  const parts: Buffer[] = []
  const dir: Buffer[] = []
  let at = 0
  for (const [name, text] of Object.entries(files)) {
    const [n, d] = [Buffer.from(name), Buffer.from(text)]
    const crc = crc32(d)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(d.length, 18)
    local.writeUInt32LE(d.length, 22)
    local.writeUInt16LE(n.length, 26)
    const entry = Buffer.alloc(46)
    entry.writeUInt32LE(0x02014b50, 0)
    entry.writeUInt16LE(20, 4)
    entry.writeUInt16LE(20, 6)
    entry.writeUInt32LE(crc, 16)
    entry.writeUInt32LE(d.length, 20)
    entry.writeUInt32LE(d.length, 24)
    entry.writeUInt16LE(n.length, 28)
    entry.writeUInt32LE(at, 42)
    parts.push(local, n, d)
    dir.push(entry, n)
    at += 30 + n.length + d.length
  }
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(dir.length / 2, 8)
  end.writeUInt16LE(dir.length / 2, 10)
  end.writeUInt32LE(dir.reduce((s, b) => s + b.length, 0), 12)
  end.writeUInt32LE(at, 16)
  return new Uint8Array(Buffer.concat([...parts, ...dir, end]))
}
