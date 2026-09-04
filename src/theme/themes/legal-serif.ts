import type { DocumentBlockStyle } from "../../domain/style.ts";
import { defineTheme, type Theme } from "../theme.ts";

/**
 * Everything the three banners share.
 *
 * No colour, because this theme has none — the three states are told apart by
 * ink and rule the way a filing tells them apart. This is the whole point of a
 * component naming what a banner *is* rather than what colour it should take.
 */
const banner: DocumentBlockStyle = {
  heightPt: 44,
  paddingPt: 14,
  valign: "center",
  fontSizePt: 12,
  weight: "bold",
  border: "000000",
  borderWidthPt: 1,
  transform: "uppercase",
};

/**
 * The theme for documents that may be read aloud in a dispute.
 *
 * Times New Roman at 12pt is less a style choice than a convention, and the
 * convention is the point: a notice set in anything else invites a question
 * about the notice rather than about its contents. Everything is black, because
 * a colour that does not survive a photocopier cannot be load-bearing.
 */
export const legalSerifTheme: Theme = defineTheme({
  id: "legal-serif",
  title: "Legal Serif",
  summary: "Times New Roman at 12pt, all black. Convention, on purpose.",
  detail:
    "Notices, terms, statements of case — anything that has to look like what " +
    "it is at a glance and survive being photocopied twice. Nothing carries " +
    "meaning by colour: headings are distinguished by weight and capitals " +
    "alone, so the document reads the same in monochrome as in full colour.",
  category: "Legal",
  tags: ["serif", "formal", "monochrome", "notice"],
  style: {
    page: {
      size: "A4",
      orientation: "portrait",
      margins: { topMm: 25.4, rightMm: 25.4, bottomMm: 25.4, leftMm: 31.75 },
    },
    typography: {
      bodyFont: "Times New Roman",
      headingFont: "Times New Roman",
      bodySizePt: 12,
      bodyLineHeight: 1.4,
      color: "000000",
    },
    palette: {
      heading: "000000",
      accent: "000000",
      muted: "3F3F3F",
      rule: "000000",
      page: "FFFFFF",
    },
    paragraph: { spacingAfterPt: 11 },
    title: {
      fontSizePt: 14,
      weight: "bold",
      spacingBeforePt: 0,
      spacingAfterPt: 16,
      transform: "uppercase",
    },
    sectionHeading: {
      fontSizePt: 12,
      weight: "bold",
      spacingBeforePt: 15,
      spacingAfterPt: 6,
      transform: "uppercase",
    },
    // Three states in one ink: ruled, filled grey, reversed out. Convention, on
    // purpose — the same reason nothing else here is coloured.
    blocks: {
      bannerPositive: { ...banner, fill: "FFFFFF", color: "000000" },
      bannerAttention: { ...banner, fill: "E5E5E5", color: "000000" },
      bannerCritical: { ...banner, fill: "000000", color: "FFFFFF" },
    },
  },
});
