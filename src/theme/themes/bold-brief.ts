import type { DocumentBlockStyle } from "../../domain/style.ts";
import { defineTheme, type Theme } from "../theme.ts";

/**
 * Everything the three banners share.
 *
 * Larger and tracked, because this theme has one page to make one point and a
 * banner is usually the point.
 */
const banner: DocumentBlockStyle = {
  color: "FFFFFF",
  heightPt: 48,
  paddingPt: 16,
  valign: "center",
  fontSizePt: 14,
  weight: "bold",
  transform: "uppercase",
  letterSpacingEm: 0.04,
};

/**
 * One page, read from across a desk.
 *
 * A large accent title over short sans-serif body text on US Letter — the
 * proportions of a briefing note handed to somebody who has ninety seconds.
 * The only shipped theme that lets the accent run at title size, because the
 * whole design assumes one page carrying one point.
 */
export const boldBriefTheme: Theme = defineTheme({
  id: "bold-brief",
  title: "Bold Brief",
  summary: "US Letter, an outsized accent title, one page and one point.",
  detail:
    "Briefing notes, one-pagers, anything summarised for somebody who will " +
    "not turn the page. The title is set large and in the accent so the " +
    "subject is legible before the document is picked up, and body text sits " +
    "at 12pt on a short measure so a paragraph is over before attention is.",
  category: "Marketing",
  tags: ["sans", "letter-size", "one-page", "accent"],
  style: {
    page: {
      size: "LETTER",
      orientation: "portrait",
      margins: { topMm: 22, rightMm: 24, bottomMm: 22, leftMm: 24 },
    },
    typography: {
      bodyFont: "Verdana",
      headingFont: "Verdana",
      bodySizePt: 12,
      bodyLineHeight: 1.45,
      color: "141414",
    },
    palette: {
      heading: "141414",
      accent: "C2185B",
      muted: "6E6E6E",
      rule: "E0E0E0",
      page: "FFFFFF",
    },
    paragraph: { spacingAfterPt: 12 },
    title: {
      fontSizePt: 30,
      weight: "bold",
      spacingBeforePt: 0,
      spacingAfterPt: 22,
      color: "C2185B",
    },
    sectionHeading: {
      fontSizePt: 13,
      weight: "bold",
      spacingBeforePt: 20,
      spacingAfterPt: 6,
      color: "141414",
    },
    // The attention tone is the theme's own accent, so the common case is the
    // colour the document already shouts in.
    blocks: {
      bannerPositive: { ...banner, fill: "00695C" },
      bannerAttention: { ...banner, fill: "C2185B" },
      bannerCritical: { ...banner, fill: "B71C1C" },
    },
  },
});
