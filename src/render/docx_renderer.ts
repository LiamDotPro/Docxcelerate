/**
 * Packing a finished document into DOCX bytes.
 *
 * @module
 */

import { Packer } from "docx";
import type { DocumentModel } from "../domain/types.ts";
import { createDocxDocument } from "./docx_document.ts";

export { createDocxDocument } from "./docx_document.ts";

/**
 * Packs a document model into `.docx` bytes.
 *
 * The counterpart to `createDocxBlob`, for writing straight to a file or a
 * response body rather than handing a blob to a browser.
 *
 * @param model The finished document.
 * @returns The file's bytes.
 */
export async function renderDocxBytes(model: DocumentModel): Promise<Uint8Array> {
  const doc = createDocxDocument(model);
  const buffer = await Packer.toBuffer(doc);

  return buffer;
}
