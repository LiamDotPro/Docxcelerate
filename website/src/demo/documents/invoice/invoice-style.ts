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
    letterSpacingEm: 0.14,
    fontSizePt: 23,
    weight: "regular",
    spacingBeforePt: 0,
    spacingAfterPt: 10,
    color: "2C3D8F",
    transform: "uppercase",
  },
  sectionHeading: {
    letterSpacingEm: 0.12,
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
      valign: "center",
      fill: "F4F6FD",
      borderSides: ["bottom"],
      border: "E3E7F5",
      paddingPt: 10,
    },
    /** A tinted box: the totals, the payment reference. */
    panel: {
      lineHeight: 1.2,
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
      // A strip, not a line of type: three pixels of navy, which is what the
      // design draws and what a depth stated in points says outright.
      heightPt: 2.25,
      spacingAfterPt: 0,
    },
    /**
     * The dark strip at the foot of every page.
     *
     * It bleeds: a bar with a white gutter either side of it is not a bar, it
     * is a box. Set on one line, because the strip is one line of small print
     * rather than a block of it.
     */
    footerBar: {
      valign: "center",
      fill: "1E2A66",
      // The design's rgba(255,255,255,0.85) over #1E2A66, composited. Not an
      // approximation — the same colour, worked out once at build time.
      color: "D5D8E4",
      bleed: true,
      // The strip's depth is padding, not leading: it holds one line of small
      // print and the design still draws it 55px deep.
      paddingPt: 15,
      fontSizePt: 7.5,
      lineHeight: 1.2,
    },
    /**
     * The last cell of the footer bar.
     *
     * The bar runs to the paper's edge; its page number stops where the design
     * stops it. Saying so on the cell beats a spacer column the document does
     * not otherwise have — and the bar's fill and centring still reach it.
     */
    footerEdge: {
      paddingSidesPt: { right: 46 },
    },
    /** The row a reader's eye stops on. */
    totalRow: {
      lineHeight: 1.2,
      fill: "1E2A66",
      color: "FFFFFF",
      weight: "bold",
      fontSizePt: 11,
    },
    /** A small capital label above a value. */
    label: {
      lineHeight: 1.2,
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
    /**
     * Every other charge row, tinted.
     *
     * Named rather than applied: the renderer counts rows as it draws them, so
     * this survives publishing, where a row does not know it is odd.
     */
    rowAlt: {
      fill: "F7F8FD",
    },
    /**
     * The covering note, held to a measure.
     *
     * Prose run across the whole text column is prose the eye loses its place
     * tracking back from. The measure narrows the column from the right, so
     * the table below it still stands exactly where it stood.
     */
    summary: {
      maxWidthMm: 158,
      fontSizePt: 9,
      lineHeight: 1.45,
    },
    /** The note under a description, and anything else set quietly. */
    chargeNote: {
      color: "5A6482",
      fontSizePt: 8.5,
      lineHeight: 1.6,
    },
    muted: {
      color: "5A6482",
      fontSizePt: 8.5,
      // Tighter than the line it sits under: it is an aside, and prose leading
      // under a description is what made a two-line charge row three deep.
      lineHeight: 1.1,
    },
    /**
     * A figure in a column of figures.
     *
     * Proportional digits are each their own width, so a column of them lines
     * up on nothing. Consolas gives every digit the same width and the column
     * reads as a column — which is the point, the face is only how it is got.
     */
    money: {
      font: "Consolas",
      lineHeight: 1.2,
      borderSides: [],
    },
    /**
     * A stacked cell: an address, a name over a value.
     *
     * A charge row's leading is set for a description with a note under it;
     * a five-line address on the same setting is airier than the design draws
     * it, and the two are not the same kind of thing.
     */
    addressCell: {
      lineHeight: 1.45,
    },
    /** A band cell: a table row's leading. The band says the rest. */
    bandCell: {
      lineHeight: 1.2,
    },
    /**
     * A row of a table, set as a row rather than as prose.
     *
     * Body leading is for paragraphs a reader travels through; a table is
     * scanned down instead. The number is what puts a charge row at the 54px
     * the design draws it at, now that a leading means the same thing in both
     * engines.
     */
    lineItem: {
      lineHeight: 1.62,
      // Stripes instead of rules: a table wearing both is wearing belt and
      // braces. An empty side list says so — it is a decision, not an omission.
      borderSides: [],
    },
  },
};
