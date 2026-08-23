import type { DocumentStyle } from "docxcelerate";

/**
 * Fernhill's house style: navy, tight, and set for figures.
 *
 * An invoice is read in two passes — what is owed, then what for — so the
 * headings are small capitals a reader skims past rather than titles that stop
 * them, and the body is set a little tighter than a letter's. The navy is the
 * sender's, not the site's: a document carries its own brand.
 *
 * The blocks below are the other half of that. A component says a node is a
 * `band` or a `badge`; this is where the theme decides what those look like, so
 * no colour is ever written into a node.
 */
export const invoiceStyle: DocumentStyle = {
  preset: "fernhill-invoice",
  page: {
    size: "A4",
    orientation: "portrait",
    margins: { topMm: 16, rightMm: 16, bottomMm: 16, leftMm: 16 },
  },
  typography: {
    bodyFont: "Aptos",
    headingFont: "Aptos",
    bodySizePt: 10,
    bodyLineHeight: 1.45,
    color: "1C2340",
  },
  palette: {
    heading: "2C3D8F",
    accent: "2C3D8F",
    muted: "5A6482",
    rule: "D9DDEB",
    page: "FFFFFF",
  },
  paragraph: { spacingAfterPt: 6 },
  // The letterhead carries the wordmark beside the reference, so a second
  // title above it would be the document naming itself twice.
  showTitle: false,
  title: {
    fontSizePt: 23,
    weight: "regular",
    spacingBeforePt: 0,
    spacingAfterPt: 10,
    color: "2C3D8F",
    transform: "uppercase",
  },
  sectionHeading: {
    fontSizePt: 7.5,
    weight: "bold",
    spacingBeforePt: 12,
    spacingAfterPt: 4,
    color: "2C3D8F",
    transform: "uppercase",
  },
  blocks: {
    /**
     * The tinted strip of dates under the letterhead.
     *
     * It stands on the text column rather than bleeding into the margins. A
     * table that reaches past them carries its first label out to the paper's
     * edge and leaves the status pill ending where no other column does —
     * and Word will not indent a table out of its margins in any case, so a
     * band that bled would be a band only the preview could draw.
     */
    band: {
      fill: "F4F6FD",
      borderSides: ["bottom"],
      border: "E3E7F5",
      paddingPt: 10,
    },
    /** A tinted box: the totals, the payment reference. */
    panel: {
      fill: "F4F6FD",
      paddingPt: 9,
    },
    /** An outlined box that is not tinted — the scan-to-pay card. */
    card: {
      border: "D9DDEB",
      paddingPt: 12,
    },
    /** The status pill, when there is still something to pay. */
    badge: {
      fill: "FBF0DC",
      border: "E5C78A",
      color: "8A5A06",
      paddingPt: 5,
      fontSizePt: 7,
      weight: "bold",
      transform: "uppercase",
      letterSpacingEm: 0.1,
    },
    /** The same pill, once it is settled. */
    "badge-done": {
      fill: "2C3D8F",
      border: "2C3D8F",
      color: "FFFFFF",
      paddingPt: 5,
      fontSizePt: 7,
      weight: "bold",
      transform: "uppercase",
      letterSpacingEm: 0.1,
    },
    /** The line under the letterhead, drawn edge to edge of the paper. */
    rule: {
      fill: "2C3D8F",
      bleed: true,
      paddingPt: 1.5,
    },
    /** The dark strip at the foot of every page. */
    footerBar: {
      fill: "1E2A66",
      color: "FFFFFF",
      paddingPt: 9,
      fontSizePt: 7.5,
    },
    /** The row a reader's eye stops on. */
    totalRow: {
      fill: "1E2A66",
      color: "FFFFFF",
      weight: "bold",
      fontSizePt: 11,
    },
    /** A small capital label above a value. */
    label: {
      color: "2C3D8F",
      fontSizePt: 7,
      weight: "bold",
      transform: "uppercase",
      letterSpacingEm: 0.12,
    },
    /** The sender's name at the top of the page. */
    senderName: {
      fontSizePt: 13.5,
      weight: "bold",
      color: "1C2340",
    },
    /** The word INVOICE, set light and opened right up. */
    wordmark: {
      fontSizePt: 23,
      color: "2C3D8F",
      transform: "uppercase",
      letterSpacingEm: 0.14,
    },
    /** The invoice number under it. */
    reference: {
      fontSizePt: 8.5,
      color: "5A6482",
    },
    /** The note under a description, and anything else set quietly. */
    muted: {
      color: "5A6482",
      fontSizePt: 8.5,
    },
  },
};
