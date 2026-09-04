import type { DocumentBlockStyle } from "../../domain/style.ts";
import { defineTheme, type Theme } from "../theme.ts";

/**
 * Everything the three banners share, so only the colour differs between them.
 *
 * The depth is here rather than left to the node because a banner's whole
 * point is being the same size whatever it says, and `valign` is what stops
 * one line of text sitting against the top edge of a 44pt box.
 */
const banner: DocumentBlockStyle = {
  color: "FFFFFF",
  heightPt: 44,
  paddingPt: 14,
  valign: "center",
  fontSizePt: 12,
  weight: "bold",
};

/**
 * The theme a document falls back to when it names none.
 *
 * Deliberately unremarkable: A4, one-inch margins, Aptos over Cambria. It is
 * the theme you should be able to send to anybody without thinking about it,
 * and the one every other theme is a departure from.
 */
export const cleanMinimalTheme: Theme = defineTheme({
  id: "clean-minimal",
  title: "Clean Minimal",
  summary: "A4, one-inch margins, Aptos over Cambria. The default.",
  detail:
    "What a document looks like when nobody has decided what it should look " +
    "like — which is the right answer more often than it sounds. Nothing here " +
    "competes with the words: one weight of grey, no rules, no accent, and " +
    "spacing loose enough to read on a screen without turning two pages into " +
    "three.",
  category: "Correspondence",
  tags: ["default", "neutral", "sans"],
  style: {
    page: {
      size: "A4",
      orientation: "portrait",
      margins: { topMm: 25.4, rightMm: 25.4, bottomMm: 25.4, leftMm: 25.4 },
    },
    typography: {
      bodyFont: "Aptos",
      headingFont: "Cambria",
      bodySizePt: 11,
      bodyLineHeight: 1.35,
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
    // A banner is the one place this theme raises its voice, because a notice
    // nobody notices is not a notice. The three are named for the state they
    // report rather than the colour they are, so a document that says
    // `bannerCritical` keeps meaning that under a theme that draws it in black.
    blocks: {
      bannerPositive: { ...banner, fill: "15803D" },
      bannerAttention: { ...banner, fill: "B45309" },
      bannerCritical: { ...banner, fill: "B91C1C" },
    },
  },
});
