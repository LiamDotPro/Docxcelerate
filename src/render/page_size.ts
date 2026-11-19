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

/** Word's own distance for a page border measured from the edge of the page. */
const DEFAULT_PAGE_BORDER_OFFSET_PT = 24;

/**
 * How far the page border stands in from the paper, in the whole points `w:space` holds.
 * Checked because Word caps a border's spacing at 31pt, and a frame quietly drawn
 * nearer the edge than asked is worse than a build that says so.
 */
export function pageBorderOffsetPt(page: DocumentPageStyle): number {
  if (page.borderOffsetMm === undefined) return DEFAULT_PAGE_BORDER_OFFSET_PT;
  const pt = Math.round(page.borderOffsetMm * 72 / 25.4);
  if (!(pt >= 0 && pt <= 31)) {
    throw new RangeError(`page.borderOffsetMm must be from 0 to 11mm, not ${page.borderOffsetMm}`);
  }
  return pt;
}

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
