import type { DocumentStyle } from "../domain/types.ts";

/**
 * The style a document falls back to when it declares none.
 *
 * @module
 */

/**
 * A plain, readable style: A4 portrait, one-inch margins, and Aptos at 11pt
 * over Cambria headings.
 *
 * This is what a renderer uses when a document carries no style of its own.
 */
export const cleanMinimalDocumentStyle: DocumentStyle = {
  preset: "clean-minimal",
  page: {
    size: "A4",
    orientation: "portrait",
    margins: {
      topMm: 25.4,
      rightMm: 25.4,
      bottomMm: 25.4,
      leftMm: 25.4,
    },
  },
  typography: {
    bodyFont: "Aptos",
    headingFont: "Cambria",
    bodySizePt: 11,
    bodyLineHeight: 1.35,
    color: "111827",
  },
  paragraph: {
    spacingAfterPt: 10,
  },
  title: {
    fontSizePt: 20,
    weight: "bold",
    spacingBeforePt: 0,
    spacingAfterPt: 18,
  },
  sectionHeading: {
    fontSizePt: 12,
    weight: "bold",
    spacingBeforePt: 16,
    spacingAfterPt: 7,
  },
};
