import { once } from 'events'

export interface ZipWriterEntryOptions {
  mtime?: Date
}

interface ZipArchiveEntry {
  name: string
  crc32: number
  size: number
  offset: number
  dosDate: number
  dosTime: number
}

const CRC32_TABLE = buildCrc32Table()
const ZIP_LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50
const ZIP_CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50
const ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50
const ZIP_VERSION = 20
const ZIP_UTF8_FLAG = 0x0800

function buildCrc32Table(): Uint32Array {
  const table = new Uint32Array(256)
  for (let index = 0; index < 256; index += 1) {
    let crc = index
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc & 1) !== 0 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1
    }
    table[index] = crc >>> 0
  }
  return table
}

function crc32(buffer: Buffer): number {
  let crc = 0xffffffff
  for (let index = 0; index < buffer.length; index += 1) {
    crc = CRC32_TABLE[(crc ^ buffer[index]) & 0xff] ^ (crc >>> 8)
  }
  return (~crc) >>> 0
}

function toDosDateTime(value: Date): { date: number; time: number } {
  const date = value instanceof Date && !Number.isNaN(value.getTime()) ? value : new Date()
  const year = Math.max(1980, date.getFullYear())
  const month = date.getMonth() + 1
  const day = date.getDate()
  const hours = date.getHours()
  const minutes = date.getMinutes()
  const seconds = Math.floor(date.getSeconds() / 2)

  return {
    date: ((year - 1980) << 9) | (month << 5) | day,
    time: (hours << 11) | (minutes << 5) | seconds
  }
}

export function normalizeZipEntryName(name: string): string {
  return String(name || '')
    .replace(/\\/g, '/')
    .replace(/^\/+/, '')
    .split('/')
    .map((segment) => segment.trim())
    .filter(Boolean)
    .join('/')
}

export function estimateZipEntrySize(name: string, contentLength: number): number {
  const normalizedName = normalizeZipEntryName(name)
  return 76 + Buffer.byteLength(normalizedName, 'utf8') * 2 + Math.max(0, contentLength)
}

function toBuffer(value: Buffer | string): Buffer {
  return Buffer.isBuffer(value) ? value : Buffer.from(String(value), 'utf8')
}

async function writeBuffer(
  stream: NodeJS.WritableStream,
  buffer: Buffer,
  state: { offset: number }
): Promise<void> {
  if (!buffer.length) {
    return
  }
  if (!stream.write(buffer)) {
    await once(stream, 'drain')
  }
  state.offset += buffer.length
}

function writeLocalFileHeader(entry: ZipArchiveEntry): Buffer {
  const nameBuffer = Buffer.from(entry.name, 'utf8')
  const header = Buffer.alloc(30)
  header.writeUInt32LE(ZIP_LOCAL_FILE_HEADER_SIGNATURE, 0)
  header.writeUInt16LE(ZIP_VERSION, 4)
  header.writeUInt16LE(ZIP_UTF8_FLAG, 6)
  header.writeUInt16LE(0, 8)
  header.writeUInt16LE(entry.dosTime, 10)
  header.writeUInt16LE(entry.dosDate, 12)
  header.writeUInt32LE(entry.crc32, 14)
  header.writeUInt32LE(entry.size, 18)
  header.writeUInt32LE(entry.size, 22)
  header.writeUInt16LE(nameBuffer.length, 26)
  header.writeUInt16LE(0, 28)
  return Buffer.concat([header, nameBuffer])
}

function writeCentralDirectoryEntry(entry: ZipArchiveEntry): Buffer {
  const nameBuffer = Buffer.from(entry.name, 'utf8')
  const header = Buffer.alloc(46)
  header.writeUInt32LE(ZIP_CENTRAL_DIRECTORY_SIGNATURE, 0)
  header.writeUInt16LE((3 << 8) | ZIP_VERSION, 4)
  header.writeUInt16LE(ZIP_VERSION, 6)
  header.writeUInt16LE(ZIP_UTF8_FLAG, 8)
  header.writeUInt16LE(0, 10)
  header.writeUInt16LE(entry.dosTime, 12)
  header.writeUInt16LE(entry.dosDate, 14)
  header.writeUInt32LE(entry.crc32, 16)
  header.writeUInt32LE(entry.size, 20)
  header.writeUInt32LE(entry.size, 24)
  header.writeUInt16LE(nameBuffer.length, 28)
  header.writeUInt16LE(0, 30)
  header.writeUInt16LE(0, 32)
  header.writeUInt16LE(0, 34)
  header.writeUInt16LE(0, 36)
  header.writeUInt32LE(0, 38)
  header.writeUInt32LE(entry.offset, 42)
  return Buffer.concat([header, nameBuffer])
}

function writeEndOfCentralDirectory(
  entryCount: number,
  centralDirectorySize: number,
  centralDirectoryOffset: number
): Buffer {
  const buffer = Buffer.alloc(22)
  buffer.writeUInt32LE(ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE, 0)
  buffer.writeUInt16LE(0, 4)
  buffer.writeUInt16LE(0, 6)
  buffer.writeUInt16LE(entryCount, 8)
  buffer.writeUInt16LE(entryCount, 10)
  buffer.writeUInt32LE(centralDirectorySize, 12)
  buffer.writeUInt32LE(centralDirectoryOffset, 16)
  buffer.writeUInt16LE(0, 20)
  return buffer
}

export class ZipStreamWriter {
  private readonly entries: ZipArchiveEntry[] = []
  private readonly state = { offset: 0 }
  private finalized = false

  constructor(private readonly stream: NodeJS.WritableStream) {}

  async addFile(name: string, content: Buffer | string, options: ZipWriterEntryOptions = {}): Promise<void> {
    if (this.finalized) {
      throw new Error('ZIP archive already finalized')
    }

    const entryName = normalizeZipEntryName(name)
    if (!entryName) {
      throw new Error('ZIP entry name is required')
    }

    const buffer = toBuffer(content)
    const mtime = options.mtime instanceof Date ? options.mtime : new Date()
    const { date, time } = toDosDateTime(mtime)
    const entry: ZipArchiveEntry = {
      name: entryName,
      crc32: crc32(buffer),
      size: buffer.length,
      offset: this.state.offset,
      dosDate: date,
      dosTime: time
    }

    const header = writeLocalFileHeader(entry)
    await writeBuffer(this.stream, header, this.state)
    await writeBuffer(this.stream, buffer, this.state)
    this.entries.push(entry)
  }

  async addText(name: string, content: string, options: ZipWriterEntryOptions = {}): Promise<void> {
    await this.addFile(name, Buffer.from(content, 'utf8'), options)
  }

  async finalize(): Promise<void> {
    if (this.finalized) {
      return
    }

    const centralDirectoryOffset = this.state.offset
    let centralDirectorySize = 0
    for (const entry of this.entries) {
      const directoryEntry = writeCentralDirectoryEntry(entry)
      centralDirectorySize += directoryEntry.length
      await writeBuffer(this.stream, directoryEntry, this.state)
    }

    const endRecord = writeEndOfCentralDirectory(
      this.entries.length,
      centralDirectorySize,
      centralDirectoryOffset
    )
    await writeBuffer(this.stream, endRecord, this.state)
    this.finalized = true
  }
}

export function createZipStreamWriter(stream: NodeJS.WritableStream): ZipStreamWriter {
  return new ZipStreamWriter(stream)
}
