import { renderLetterWebsite, type WebRenderOptions } from "../renderer/web_renderer.ts";
import type { LetterDocument } from "../domain/types.ts";

export type HtmlRenderOptions = WebRenderOptions;

export function renderLetterHtml(
  letter: LetterDocument,
  options: HtmlRenderOptions = {},
): string {
  return renderLetterWebsite(letter, options);
}
