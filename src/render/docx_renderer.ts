import { Packer } from "docx";
import type { DocumentModel } from "../domain/types.ts";
import { createDocxDocument } from "./docx_document.ts";

export { createDocxDocument } from "./docx_document.ts";

export async function renderDocxBytes(model: DocumentModel): Promise<Uint8Array> {
  const doc = createDocxDocument(model);
  const buffer = await Packer.toBuffer(doc);

  return buffer;
}
