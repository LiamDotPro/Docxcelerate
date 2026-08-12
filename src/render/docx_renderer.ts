import { Packer } from "docx";
import type { LetterDocument } from "../domain/types.ts";
import { createDocxDocument } from "./docx_document.ts";

export { createDocxDocument } from "./docx_document.ts";

export async function renderDocxBytes(letter: LetterDocument): Promise<Uint8Array> {
  const doc = createDocxDocument(letter);
  const buffer = await Packer.toBuffer(doc);

  return buffer;
}
