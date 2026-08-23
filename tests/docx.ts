import { inflateRawSync } from "node:zlib";
import { createDocxBlob } from "docxcelerate/docx";
import type { DocumentModel } from "docxcelerate";

/**
 * Opening the file the framework exists to produce.
 *
 * A `.docx` is a zip of XML, so the only way to know a fill, a border or a
 * page size survived packing is to unpack it and look. Everything that used to
 * be asserted against an HTML preview is asserted here instead: there is one
 * renderer now, and this is how it is read back.
 *
 * @module
 */

/** Reads one entry out of a zip, via its central directory. */
export function entryOf(zip: Uint8Array, wanted: string): Uint8Array {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  const text = new TextDecoder("latin1").decode(zip);

  // The end-of-central-directory record says where the directory starts. It is
  // last in the file, so it is found by scanning back for its signature.
  const end = text.lastIndexOf("PK");

  if (end === -1) {
    throw new Error("not a zip");
  }

  let at = view.getUint32(end + 16, true);
  const count = view.getUint16(end + 10, true);

  for (let index = 0; index < count; index += 1) {
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const name = text.slice(at + 46, at + 46 + nameLength);

    if (name === wanted) {
      const method = view.getUint16(at + 10, true);
      const size = view.getUint32(at + 20, true);
      const offset = view.getUint32(at + 42, true);
      // The local header repeats the name and extra fields, at its own lengths.
      const localName = view.getUint16(offset + 26, true);
      const localExtra = view.getUint16(offset + 28, true);
      const start = offset + 30 + localName + localExtra;
      const data = zip.slice(start, start + size);

      return method === 0 ? data : inflateRawSync(data);
    }

    at += 46 + nameLength + extraLength + commentLength;
  }

  throw new Error(`no ${wanted} in the package`);
}

/** One part of the packed file, as text. */
export async function partXml(doc: DocumentModel, part: string): Promise<string> {
  const blob = await createDocxBlob(doc);
  const bytes = new Uint8Array(await blob.arrayBuffer());

  return new TextDecoder().decode(entryOf(bytes, part));
}

/** The body of the packed document. */
export function documentXml(doc: DocumentModel): Promise<string> {
  return partXml(doc, "word/document.xml");
}

/** Every part the packed file holds, by name. */
export async function partNames(doc: DocumentModel): Promise<string[]> {
  const blob = await createDocxBlob(doc);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const text = new TextDecoder("latin1").decode(bytes);
  const end = text.lastIndexOf("PK");
  const count = view.getUint16(end + 10, true);
  const names: string[] = [];

  let at = view.getUint32(end + 16, true);

  for (let index = 0; index < count; index += 1) {
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);

    names.push(text.slice(at + 46, at + 46 + nameLength));
    at += 46 + nameLength + extraLength + commentLength;
  }

  return names;
}

/** The text a run prints, with the XML around it taken off. */
export function textOf(xml: string): string {
  return [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1]).join("");
}
