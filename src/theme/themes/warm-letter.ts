import type { DocumentBlockStyle } from "../../domain/style.ts";
import { defineTheme, type Theme } from "../theme.ts";

/**
 * Everything the three banners share.
 *
 * The ink is the page's own warm white rather than a flat `FFFFFF`, which is
 * the difference between a banner that belongs on this paper and one that was
 * cut out and glued onto it.
 */
const banner: DocumentBlockStyle = {
  color: "FFFDF9",
  heightPt: 44,
  paddingPt: 15,
  valign: "center",
  fontSizePt: 12,
  weight: "bold",
};

/**
 * Correspondence that is meant to sound like a person wrote it.
 *
 * Georgia at 11.5pt on generous leading, wide margins, and a short measure —
 * the shape of a letter rather than of a document. The accent is warm because
 * the theme is for the letters an organisation would rather not send coldly:
 * renewals, apologies, decisions that went against somebody.
 */
export const warmLetterTheme: Theme = defineTheme({
  id: "warm-letter",
  title: "Warm Letter",
  summary: "Georgia on wide margins. Correspondence that reads like a person.",
  detail:
    "A serif face, loose leading, and margins wide enough to shorten the line " +
    "to something the eye finishes comfortably. Use it where the tone of the " +
    "document matters as much as its contents — renewals, apologies, refusals " +
    "— and where a reader should not feel they have been sent a form.",
  category: "Correspondence",
  tags: ["serif", "letter", "generous", "warm"],
  style: {
    page: {
      size: "A4",
      orientation: "portrait",
      margins: { topMm: 30, rightMm: 32, bottomMm: 28, leftMm: 32 },
    },
    typography: {
      bodyFont: "Georgia",
      headingFont: "Georgia",
      bodySizePt: 11.5,
      bodyLineHeight: 1.5,
      color: "2B2118",
    },
    palette: {
      heading: "42301F",
      accent: "A2571B",
      muted: "8A7A6A",
      rule: "E2D6C7",
      page: "FFFDF9",
    },
    paragraph: { spacingAfterPt: 12 },
    title: {
      fontSizePt: 19,
      weight: "regular",
      spacingBeforePt: 0,
      spacingAfterPt: 20,
      color: "42301F",
    },
    sectionHeading: {
      fontSizePt: 12.5,
      weight: "bold",
      spacingBeforePt: 18,
      spacingAfterPt: 8,
      color: "42301F",
    },
    // Muted rather than signal-bright: a letter that shouts is a letter nobody
    // believes. The attention tone is the theme's own accent.
    blocks: {
      bannerPositive: { ...banner, fill: "4F6F52" },
      bannerAttention: { ...banner, fill: "A2571B" },
      bannerCritical: { ...banner, fill: "8C2F1F" },
    },
  },
});
