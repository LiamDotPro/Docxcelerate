import { renderDocumentWebsite, type WebRenderOptions } from "../renderer/web_renderer.ts";
import type { DocumentModel } from "../domain/types.ts";

export type HtmlRenderOptions = WebRenderOptions;

export function renderDocumentHtml(
  letter: DocumentModel,
  options: HtmlRenderOptions = {},
): string {
  return renderDocumentWebsite(letter, options);
}
