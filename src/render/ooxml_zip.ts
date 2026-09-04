/**
 * The zip a `.docx` is, read apart and put back together.
 *
 * A Word chart is not one element in `document.xml` — it is a part of its own
 * beside it, with a relationship pointing at it, a content type declaring it
 * and a workbook hanging off it. The library that lays the document out packs
 * the zip itself and has no way to be handed a part, so the package is opened
 * after it is written and the chart's parts are added to it.
 *
 * Everything here is web-standard, for the same reason
 * {@linkcode ../render/docx_packed.ts | docx_packed} is: `CompressionStream`
 * and `DecompressionStream` are in Node 20 and in every browser the preview
 * runs in, so a document packs the same way in a scaffolded workspace's tab as
 * it does on a server. There is no `node:zlib` here, because half the callers
 * have no Node to import it from.
 *
 * @module
 */

/** One file inside the package. */
export interface ZipEntry {
  /** Its path within the zip, as `/`-separated names. */
  readonly name: string;
  /** Its contents, uncompressed. */
  readonly bytes: Uint8Array;
}

/** The CRC-32 table, built once — the checksum every zip entry records. */
const CRC_TABLE: Uint32Array = (() => {
  const table = new Uint32Array(256);

  for (let index = 0; index < 256; index += 1) {
    let value = index;

    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }

    table[index] = value >>> 0;
  }

  return table;
})();

/** The signature at the head of each local file record. */
const LOCAL_SIGNATURE = 0x04034b50;
/** The signature at the head of each central directory record. */
const CENTRAL_SIGNATURE = 0x02014b50;
/** The signature at the head of the end-of-central-directory record. */
const END_SIGNATURE = 0x06054b50;
/**
 * The DOS timestamp every entry is written with: 1980-01-01 00:00:00.
 *
 * A fixed date rather than the clock, so packing the same document twice
 * produces the same bytes. A build that differs only in when it ran is a build
 * whose output cannot be compared, cached or checksummed.
 */
const FIXED_DOS_TIME = 0;
/** The DOS date for 1980-01-01, which is the earliest a zip can record. */
const FIXED_DOS_DATE = 0x0021;

/** The CRC-32 of some bytes, which is the checksum every zip entry records. */
function crc32(bytes: Uint8Array): number {
  let value = 0xffffffff;

  for (let index = 0; index < bytes.length; index += 1) {
    value = CRC_TABLE[(value ^ bytes[index]) & 0xff] ^ (value >>> 8);
  }

  return (value ^ 0xffffffff) >>> 0;
}

/** Pushes bytes through a transform stream and collects what comes out. */
async function through(
  stream: TransformStream<BufferSource, Uint8Array>,
  data: Uint8Array,
): Promise<Uint8Array> {
  const writer = stream.writable.getWriter();

  void writer.write(data as BufferSource);
  void writer.close();

  const chunks: Uint8Array[] = [];
  const reader = stream.readable.getReader();

  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
  }

  return concat(chunks);
}

/** Several runs of bytes, end to end. */
function concat(chunks: readonly Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const out = new Uint8Array(total);
  let at = 0;

  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }

  return out;
}

/**
 * Raw deflate, through the one API both a browser and Node 20 have.
 *
 * @param data The bytes to compress.
 * @returns The compressed bytes, with no zlib wrapper.
 */
export async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
  return await through(
    new CompressionStream("deflate-raw") as TransformStream<BufferSource, Uint8Array>,
    data,
  );
}

/**
 * Raw inflate, the counterpart to {@linkcode deflateRaw}.
 *
 * @param data The compressed bytes.
 * @returns What they decompress to.
 */
export async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  return await through(
    new DecompressionStream("deflate-raw") as TransformStream<BufferSource, Uint8Array>,
    data,
  );
}

/**
 * Every part of a package, in the order the zip lists them.
 *
 * The order is kept because a package is not a set: `[Content_Types].xml` is
 * expected first, and a reader that streams the zip rather than seeking its
 * directory reads the parts in the order they were written.
 *
 * @param packed The `.docx` bytes.
 * @returns One entry per part, decompressed.
 */
export async function readZipEntries(packed: Uint8Array): Promise<ZipEntry[]> {
  const view = new DataView(packed.buffer as ArrayBuffer, packed.byteOffset, packed.byteLength);
  const latin1 = new TextDecoder("latin1").decode(packed);
  const end = latin1.lastIndexOf("PK\x05\x06");

  if (end === -1) {
    throw new Error("Not a zip: no end-of-central-directory record.");
  }

  const count = view.getUint16(end + 10, true);
  const entries: ZipEntry[] = [];
  let at = view.getUint32(end + 16, true);

  for (let index = 0; index < count; index += 1) {
    if (view.getUint32(at, true) !== CENTRAL_SIGNATURE) {
      throw new Error("Not a zip: central directory record " + index + " is malformed.");
    }

    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const name = new TextDecoder().decode(packed.subarray(at + 46, at + 46 + nameLength));
    const method = view.getUint16(at + 10, true);
    const size = view.getUint32(at + 20, true);
    const offset = view.getUint32(at + 42, true);
    // The local header repeats the name and the extra field, at its own
    // lengths — which are not always the central directory's, because a writer
    // may pad one and not the other.
    const localNameLength = view.getUint16(offset + 26, true);
    const localExtraLength = view.getUint16(offset + 28, true);
    const start = offset + 30 + localNameLength + localExtraLength;
    const stored = packed.subarray(start, start + size);

    entries.push({
      name,
      bytes: method === 0 ? stored.slice() : await inflateRaw(stored),
    });

    at += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

/**
 * A package written back out, every part deflated.
 *
 * @param entries The parts, in the order they should be written.
 * @returns The zip's bytes.
 */
export async function writeZipEntries(entries: readonly ZipEntry[]): Promise<Uint8Array> {
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const name = new TextEncoder().encode(entry.name);
    const compressed = await deflateRaw(entry.bytes);
    const checksum = crc32(entry.bytes);

    const local = new Uint8Array(30 + name.length);
    const localView = new DataView(local.buffer);

    localView.setUint32(0, LOCAL_SIGNATURE, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(8, 8, true);
    localView.setUint16(10, FIXED_DOS_TIME, true);
    localView.setUint16(12, FIXED_DOS_DATE, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, compressed.length, true);
    localView.setUint32(22, entry.bytes.length, true);
    localView.setUint16(26, name.length, true);
    local.set(name, 30);

    const central = new Uint8Array(46 + name.length);
    const centralView = new DataView(central.buffer);

    centralView.setUint32(0, CENTRAL_SIGNATURE, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(10, 8, true);
    centralView.setUint16(12, FIXED_DOS_TIME, true);
    centralView.setUint16(14, FIXED_DOS_DATE, true);
    centralView.setUint32(16, checksum, true);
    centralView.setUint32(20, compressed.length, true);
    centralView.setUint32(24, entry.bytes.length, true);
    centralView.setUint16(28, name.length, true);
    centralView.setUint32(42, offset, true);
    central.set(name, 46);

    locals.push(local, compressed);
    centrals.push(central);
    offset += local.length + compressed.length;
  }

  const directorySize = centrals.reduce((sum, record) => sum + record.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);

  endView.setUint32(0, END_SIGNATURE, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, directorySize, true);
  endView.setUint32(16, offset, true);

  return concat([...locals, ...centrals, end]);
}
