/**
 * Paper sizes and the widths worked out from them. The renderer and the
 * preview share this so a page measures the same wherever it is asked.
 *
 * @module
 */
import type { DocumentPageStyle } from "../domain/types.ts";

/** Each named paper, upright, in millimetres. */
const PAPER_MM: Record<string, { widthMm: number; heightMm: number }> = {
  A3: { widthMm: 297, heightMm: 420 },
  A4: { widthMm: 210, heightMm: 297 },
  A5: { widthMm: 148, heightMm: 210 },
  LETTER: { widthMm: 215.9, heightMm: 279.4 },
  LEGAL: { widthMm: 215.9, heightMm: 355.6 },
};

/** The sheet upright, as `w:pgSz` is given it. `docx` turns it for a landscape page itself. */
export function paperMm(page: DocumentPageStyle): { widthMm: number; heightMm: number } {
  return typeof page.size === "object" ? page.size : PAPER_MM[page.size] ?? PAPER_MM.A4;
}

/** How wide the page stands once it is turned, in mm. Bleeds are cut from this. */
export function pageWidthMm(page: DocumentPageStyle): number {
  const paper = paperMm(page);
  return page.orientation === "landscape" ? paper.heightMm : paper.widthMm;
}

/** The text column: the page as turned, less its side margins, in mm. */
export function columnWidthMm(page: DocumentPageStyle): number {
  return pageWidthMm(page) - page.margins.leftMm - page.margins.rightMm;
}
