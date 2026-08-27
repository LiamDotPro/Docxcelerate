/**
 * The style paragraph cases are set in.
 *
 * Deliberately plain and deliberately round: 20mm margins, 11pt Aptos, 10pt
 * after each paragraph, 1.4 leading. Every number here is one a case may have
 * to assert against, so each is chosen to survive the trip into twips without
 * rounding — 10pt is 200 twips exactly, 20mm is 1134.
 *
 * The title is suppressed. It is not that a title is wrong; it is that a
 * printed one would be body paragraph zero in every case, and a paragraph case
 * that has to count past its own furniture is a case that will be read wrongly.
 *
 * @module
 */

import type { DocumentStyle } from "docxcelerate";

/** The baseline: what a paragraph looks like when a case says nothing else. */
export const caseStyle: DocumentStyle = {
  preset: "conformance",
  page: {
    size: "A4",
    orientation: "portrait",
    margins: { topMm: 20, rightMm: 20, bottomMm: 20, leftMm: 20 },
  },
  typography: {
    bodyFont: "Aptos",
    headingFont: "Aptos",
    bodySizePt: 11,
    bodyLineHeight: 1.4,
    color: "111827",
  },
  palette: {
    heading: "111827",
    accent: "2F5FBD",
    muted: "6B7280",
    rule: "D1D5DB",
    page: "FFFFFF",
  },
  paragraph: { spacingAfterPt: 10 },
  title: { fontSizePt: 20, weight: "bold", spacingBeforePt: 0, spacingAfterPt: 18 },
  sectionHeading: { fontSizePt: 12, weight: "bold", spacingBeforePt: 16, spacingAfterPt: 7 },
  showTitle: false,
};

/** The same, with the blocks a case needs. Blocks merge; nothing else changes. */
export function withBlocks(blocks: DocumentStyle["blocks"]): DocumentStyle {
  return { ...caseStyle, blocks: { ...caseStyle.blocks, ...blocks } };
}

/** The text column's width, in millimetres — A4 less the two margins. */
export const COLUMN_MM = 210 - 20 - 20;

/** A millimetre in twips, which is the unit the packed file counts indents in. */
export const TWIPS_PER_MM = (72 / 25.4) * 20;

/** A point in twips. */
export const TWIPS_PER_PT = 20;
