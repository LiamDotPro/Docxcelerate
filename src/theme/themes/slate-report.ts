import type { DocumentBlockStyle } from "../../domain/style.ts";
import { defineTheme, type Theme } from "../theme.ts";

/**
 * Everything the three banners share.
 *
 * Set in the small tracked capitals this theme heads its sections with, so a
 * banner reads as part of the report rather than something pasted onto it.
 */
const banner: DocumentBlockStyle = {
  color: "FFFFFF",
  heightPt: 44,
  paddingPt: 12,
  valign: "center",
  fontSizePt: 10,
  weight: "bold",
  transform: "uppercase",
  letterSpacingEm: 0.08,
};

/**
 * The quarterly-report theme: dense, sober, and built to be skimmed.
 *
 * Tighter body spacing than a letter, because a report is read in columns of
 * attention rather than start to finish, and headings set in small capitals so
 * that a reader scanning for one section finds it without a rule to help them.
 */
export const slateReportTheme: Theme = defineTheme({
  id: "slate-report",
  title: "Slate Report",
  summary: "Dense, sober, built to be skimmed. Small-cap headings on slate.",
  detail:
    "For documents that are consulted rather than read: performance packs, " +
    "board papers, anything with more sections than a reader intends to " +
    "finish. Margins are narrower and leading tighter than a letter's, which " +
    "buys about a fifth more on the page, and the headings carry the " +
    "navigation — set small, in capitals, in a colder ink than the body.",
  category: "Report",
  tags: ["sans", "dense", "corporate", "skimmable"],
  style: {
    page: {
      size: "A4",
      orientation: "portrait",
      margins: { topMm: 20, rightMm: 18, bottomMm: 20, leftMm: 18 },
    },
    typography: {
      bodyFont: "Calibri",
      headingFont: "Calibri",
      bodySizePt: 10.5,
      bodyLineHeight: 1.25,
      color: "1F2933",
    },
    palette: {
      heading: "334155",
      accent: "0F766E",
      muted: "64748B",
      rule: "CBD5E1",
      page: "FFFFFF",
    },
    paragraph: { spacingAfterPt: 7 },
    title: {
      fontSizePt: 22,
      weight: "bold",
      spacingBeforePt: 0,
      spacingAfterPt: 14,
      color: "0F172A",
    },
    sectionHeading: {
      fontSizePt: 10,
      weight: "bold",
      spacingBeforePt: 14,
      spacingAfterPt: 5,
      color: "0F766E",
      transform: "uppercase",
    },
    // Named for the state, not the colour. The teal is the theme's own accent,
    // so the ordinary case of a banner is the report's own voice.
    blocks: {
      bannerPositive: { ...banner, fill: "0F766E" },
      bannerAttention: { ...banner, fill: "92400E" },
      bannerCritical: { ...banner, fill: "9F1239" },
    },
  },
});
