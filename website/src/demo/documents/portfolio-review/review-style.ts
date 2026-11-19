import type { DocumentBlockStyle, DocumentStyle } from "docxcelerate";

/**
 * Harbourside's house style: a deep green, Georgia for anything a reader
 * stops on, Arial for everything they scan.
 *
 * Every size, tracking and colour here is the canvas's own, read off its CSS
 * rather than eyeballed; a block's comment says where it is used when the
 * name alone does not. Colours are hex without the `#`, as Word writes them.
 */

/** What the canvas's text inherits: its paper ink, not the darker brand ink. */
const BODY = "1F2933";
const INK = "16211C";
const MUTED = "5C6B64";
const DEEP = "0E3B2D";
const GREEN = "17624A";
const MID = "2F8A6B";
const PALE = "8AC5AE";
const RULE = "C8D3CE";
const HAIR = "DFE5E2";
const TRACK = "F7F9F8";
const SERIF = "Georgia";
/**
 * The leading of anything the canvas does not set a leading on: its design
 * system sets the body at 1.7, and every run inherits it at its own size.
 */
const LEADING = 1.7;

/** 16mm, the text column's inset, in points — where a bleeding band's text starts. */
const INSET_PT = 45.35;

/** A tracked capital label: CONTENTS, PROPOSED TRADES, VALUE AT 30 SEPTEMBER. */
const label = (sizePt: number, tracking: number, color = BODY): DocumentBlockStyle => ({
  fontSizePt: sizePt,
  weight: "bold",
  transform: "uppercase",
  letterSpacingEm: tracking,
  color,
  lineHeight: LEADING,
});

/** A Georgia figure, bold, set solid. */
const figure = (sizePt: number, color = DEEP): DocumentBlockStyle => ({
  font: SERIF,
  fontSizePt: sizePt,
  weight: "bold",
  color,
  lineHeight: 1.0,
});

/** A plain run of Arial at a size and colour. */
const text = (sizePt: number, color = BODY, lineHeight = LEADING): DocumentBlockStyle => ({
  fontSizePt: sizePt,
  color,
  lineHeight,
});

/**
 * A bar segment: a filled cell with nothing in it.
 *
 * Word will not leave a cell empty, so it holds one paragraph — set a point
 * deep, or a segment three points tall would be stretched to the depth of a
 * line of body text by the preview, which grows a row its content overflows.
 */
const EMPTY: DocumentBlockStyle = {
  borderSides: [],
  paddingPt: 0,
  // Sides as well: a table's padding by side outranks a cell's single figure.
  paddingSidesPt: { top: 0, bottom: 0, left: 0, right: 0 },
  fontSizePt: 1,
  lineHeight: 1,
};
const fill = (color: string): DocumentBlockStyle => ({ ...EMPTY, fill: color });

export const reviewStyle: DocumentStyle = {
  preset: "harbourside-review",
  page: {
    size: "A4",
    orientation: "portrait",
    // The running band is the top of every page; the body starts under it.
    margins: { topMm: 12, rightMm: 16, bottomMm: 14, leftMm: 16 },
    headerMm: 0,
    footerMm: 8,
  },
  typography: {
    bodyFont: "Arial",
    headingFont: SERIF,
    bodySizePt: 9,
    bodyLineHeight: LEADING,
    color: BODY,
  },
  palette: {
    heading: BODY,
    accent: GREEN,
    muted: MUTED,
    rule: RULE,
    page: "FFFFFF",
    series: [GREEN, RULE, MID, PALE],
  },
  paragraph: { spacingAfterPt: 0 },
  showTitle: false,
  title: { fontSizePt: 19, weight: "bold", spacingBeforePt: 0, spacingAfterPt: 0 },
  sectionHeading: {
    fontSizePt: 7,
    weight: "bold",
    spacingBeforePt: 16,
    spacingAfterPt: 7,
    transform: "uppercase",
    letterSpacingEm: 0.18,
    color: INK,
    keepWithNext: true,
  },
  blocks: {
    // ── Running furniture ───────────────────────────────────────────────
    /** The dark band across the top of every page after the cover. */
    headBand: {
      fill: DEEP,
      color: "FFFFFF",
      bleed: true,
      fontSizePt: 7,
      lineHeight: LEADING,
      borderSides: [],
      paddingSidesPt: { top: 7, bottom: 7, left: 0, right: 0 },
      valign: "center",
    },
    // A block's sides are one setting: naming the left alone would drop the
    // band's top and bottom, so all four are said.
    headName: { weight: "bold", letterSpacingEm: 0.2, paddingSidesPt: { top: 7, bottom: 7, left: INSET_PT, right: 0 } },
    headMeta: { color: "B9DDCE", paddingSidesPt: { top: 7, bottom: 7, left: 0, right: INSET_PT } },
    /** The 3pt green line under the band, and under the cover. */
    headRule: { fill: MID, bleed: true, heightPt: 3, spacingAfterPt: 0 },
    /** The footer: small print over a hairline. */
    foot: {
      ...text(6.5, MUTED),
      borderSides: ["top"],
      border: RULE,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 4, bottom: 0, left: 0, right: 0 },
    },
    footPage: { weight: "bold", letterSpacingEm: 0.12, align: "right", transform: "uppercase" },

    // ── The cover ───────────────────────────────────────────────────────
    /** The paragraph the photograph is anchored in: there, but a point deep. */
    anchor: { heightPt: 1, spacingAfterPt: 0 },
    /** The words on the photograph: white, no rules, no padding of their own. */
    coverBand: { color: "FFFFFF", borderSides: [], paddingPt: 0, valign: "center" },
    coverMark: {
      ...figure(10, "FFFFFF"),
      align: "center",
      borderSides: ["top", "right", "bottom", "left"],
      border: PALE,
      borderWidthPt: 1,
      paddingPt: 0,
    },
    coverFirm: { ...label(8.5, 0.2, "FFFFFF") },
    coverPrivate: { ...label(6.5, 0.2, PALE) },
    /** The title block, standing on the foot of the photograph. */
    // The 10mm under the title block is a paragraph, not padding: Word adds a
    // cell's bottom margin to a row of exact height, where the preview keeps
    // it inside, and the whole cover slid 31px down the page in Word.
    coverFoot: { valign: "bottom", paddingSidesPt: { top: 0, bottom: 0, left: 0, right: 0 } },
    coverFootGap: { heightPt: 28.35 },
    coverKicker: { ...label(7.5, 0.2, PALE), spacingAfterPt: 5 },
    coverTitle: { ...figure(42, "FFFFFF"), spacingAfterPt: 5 },
    coverDates: text(10, "B9DDCE"),
    /** The 11mm the canvas leaves above the rule under the dates. */
    coverGap: { heightPt: 31.2 },
    coverFor: { borderSides: [], paddingPt: 0, valign: "bottom" },
    coverForCell: {
      borderSides: ["top"],
      border: "40574F",
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 19.8, bottom: 0, left: 0, right: 0 },
      valign: "bottom",
    },
    coverLabel: { ...label(6.5, 0.2, PALE), spacingAfterPt: 3 },
    coverLabelValue: { ...label(6.5, 0.2, PALE), spacingAfterPt: 2 },
    coverName: { ...figure(20, "FFFFFF"), lineHeight: 1.1, spacingAfterPt: 3 },
    coverSub: text(8.5, "B9DDCE"),
    coverValue: { ...figure(44, "FFFFFF"), spacingAfterPt: 2 },
    coverChange: text(9, "B9DDCE"),
    /** The decision and the headline, under the photograph. */
    coverDecision: {
      borderSides: ["bottom"],
      border: RULE,
      borderWidthPt: 0.75,
      paddingPt: 0,
    },
    coverHeadlineCell: { borderSides: [], paddingSidesPt: { top: 9, bottom: 9, left: 14, right: 0 }, valign: "center" },
    coverHeadline: { font: SERIF, fontSizePt: 12.5, lineHeight: 1.38, color: INK },
    /** The two-column contents list. */
    contents: {
      ...text(9),
      borderSides: ["bottom"],
      border: HAIR,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 3.5, bottom: 3.5, left: 0, right: 0 },
    },
    contentsNumber: { font: SERIF, weight: "bold", color: GREEN },
    contentsPage: { color: MUTED },
    contentsGap: { borderSides: [] },
    /** Who prepared the review, above the cover's footer. */
    signoff: { ...text(7.5, MUTED, 1.5), borderSides: [], paddingPt: 0, valign: "top" },
    signoffGap: { heightPt: 14 },

    // ── The letter ──────────────────────────────────────────────────────
    letterHead: { ...text(9.5, BODY, 1.5), borderSides: [], paddingSidesPt: { top: 28.9, bottom: 0, left: 0, right: 0 }, valign: "top" },
    letterDate: { color: MUTED },
    greeting: { ...text(10.5), spacingBeforePt: 34 },
    letterTitle: { ...figure(18, INK), lineHeight: LEADING, spacingBeforePt: 8, spacingAfterPt: 10 },
    letterOpen: { font: SERIF, fontSizePt: 10.5, lineHeight: 1.55, color: INK, spacingAfterPt: 8 },
    letterBody: {
      font: SERIF,
      fontSizePt: 10.5,
      lineHeight: 1.55,
      color: INK,
      spacingAfterPt: 8,
      dropCap: { lines: 2, variant: "letterCap" },
    },
    /** The letter's first letter, dropped two lines in the accent. */
    letterCap: { font: SERIF, fontSizePt: 36, weight: "bold", color: GREEN, paddingSidesPt: { right: 5 } },
    closing: { font: SERIF, fontSizePt: 10.5, lineHeight: LEADING, spacingBeforePt: 6, spacingAfterPt: 2 },
    signature: { font: "Brush Script MT", fontSizePt: 24, lineHeight: 1.2, color: DEEP, spacingAfterPt: 2 },
    signatory: { ...text(9), weight: "bold", spacingAfterPt: 2 },
    signatoryRole: text(9, MUTED),
    contact: {
      ...text(8.5, INK, 1.45),
      borderSides: ["top"],
      border: DEEP,
      borderWidthPt: 1.5,
      paddingSidesPt: { top: 8, bottom: 0, left: 0, right: 0 },
      valign: "top",
    },
    contactLabel: { ...label(6.5, 0.18, MUTED), spacingAfterPt: 2 },

    // ── Section openings ────────────────────────────────────────────────
    /** "1  Summary" with its standfirst, over a dark rule. */
    opening: {
      borderSides: ["bottom"],
      border: DEEP,
      borderWidthPt: 1.5,
      // The band and its rule end 28.9pt down and the canvas opens 12pt under
      // them; the margin already stands the body 34pt down, so 6.9pt is left.
      paddingSidesPt: { top: 6.9, bottom: 7, left: 0, right: 0 },
      valign: "top",
    },
    openingNumber: figure(24, GREEN),
    // The canvas sets the title on the number's baseline, which stands it
    // 4.3pt lower than its own line would; a shallow line with the space put
    // above it lowers the words without lowering the standfirst under them.
    openingTitle: { ...figure(19, BODY), lineHeight: 1.22, spacingBeforePt: 1.85 },
    standfirst: { ...text(9, MUTED), spacingBeforePt: 3 },

    /** A section label: OUR RECOMMENDATION, over a hairline. */
    rubric: {
      ...label(7, 0.18),
      borderSides: ["bottom"],
      border: RULE,
      borderWidthPt: 0.75,
      paddingSidesPt: { bottom: 3 },
      spacingBeforePt: 16,
      spacingAfterPt: 7,
      keepWithNext: true,
      tabStopsMm: [{ at: 178, align: "right" }],
    },
    /** The same label where the canvas opens it 14pt down rather than 16. */
    rubricTight: {
      ...label(7, 0.18),
      borderSides: ["bottom"],
      border: RULE,
      borderWidthPt: 0.75,
      paddingSidesPt: { bottom: 3 },
      spacingBeforePt: 14,
      spacingAfterPt: 7,
      keepWithNext: true,
      tabStopsMm: [{ at: 178, align: "right" }],
    },
    rubricNote: { ...text(7, MUTED), weight: "regular", transform: "none", letterSpacingEm: 0 },

    // ── Figures ─────────────────────────────────────────────────────────
    /** One cell of a key-figures grid: a dark top rule, a label, a figure, a note. */
    keyFigure: {
      borderSides: ["top"],
      border: DEEP,
      borderWidthPt: 1.5,
      paddingSidesPt: { top: 6, bottom: 0, left: 0, right: 0 },
      valign: "top",
    },
    keyGap: { borderSides: [], paddingPt: 0 },
    keyLabel: { ...label(6.5, 0.16, MUTED), spacingAfterPt: 3 },
    keyValue: { ...figure(24), spacingAfterPt: 2.5 },
    keyNote: text(8, MUTED),
    keyRise: text(8, GREEN),

    // ── The decision ────────────────────────────────────────────────────
    /** The green stamp: DECISION / Rebalance / recommended this quarter. */
    stamp: {
      fill: GREEN,
      color: "FFFFFF",
      // The canvas pads it 10pt; Word sets "Rebalance" a hair wider than the
      // browser does, and at 10pt it breaks inside the word.
      paddingSidesPt: { top: 9, bottom: 9, left: 7, right: 5 },
      valign: "center",
      borderSides: [],
      standIn: "stampPending",
    },
    /** The same stamp in a proof, before the engine has decided: outlined, not filled. */
    stampPending: {
      fill: "FFFFFF",
      color: MUTED,
      borderSides: ["top", "right", "bottom", "left"],
      border: RULE,
      borderStyle: "dashed",
      borderWidthPt: 0.75,
    },
    // The stamp's words take their ink from the stamp, so a pending stamp's are quiet too.
    stampLabel: { fontSizePt: 6, weight: "bold", transform: "uppercase", letterSpacingEm: 0.2, lineHeight: LEADING, spacingAfterPt: 2 },
    stampWord: { font: SERIF, fontSizePt: 15, weight: "bold", lineHeight: 1.05, spacingAfterPt: 2, standInLabel: false },
    stampNote: { fontSizePt: 7.5, lineHeight: LEADING, standInLabel: false },

    /** The box the recommendation sits in: hairline sides, a dark top. */
    recBox: {
      borderSides: ["bottom"],
      border: RULE,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 9, bottom: 9, left: 12, right: 12 },
      valign: "top",
    },
    recMiddle: { borderSides: ["bottom"] },
    recHead: { ...figure(11.5, INK), lineHeight: 1.25, spacingAfterPt: 3 },
    recWhy: { ...text(8.5, INK, 1.45), standInLabel: false },
    trades: {
      fill: "F7F9F8",
      borderSides: ["left", "right", "bottom"],
      border: RULE,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 6, bottom: 6, left: 10, right: 10 },
      valign: "center",
    },
    tradesLabel: { ...label(6, 0.18, MUTED), spacingAfterPt: 2 },
    trade: { ...text(8), paddingSidesPt: { top: 2.5, bottom: 2.5, left: 0, right: 0 }, borderSides: ["bottom"], border: HAIR, borderWidthPt: 0.75 },
    tradeSide: { weight: "bold", color: GREEN, transform: "uppercase" },
    tradeAmount: { weight: "bold", align: "right" },

    // ── Steps and points ────────────────────────────────────────────────
    stepNumber: { ...figure(18, GREEN), paddingSidesPt: { top: 0, bottom: 0, left: 0, right: 0 }, borderSides: [] },
    stepTitle: { ...text(9), weight: "bold", spacingAfterPt: 2 },
    stepBody: text(8.5, MUTED, 1.45),
    stepCell: { borderSides: [], paddingSidesPt: { top: 0, bottom: 0, left: 0, right: 0 }, valign: "top" },
    /** One of the quarter's points: a numbered line over a hairline. */
    point: {
      ...text(10, BODY, 1.45),
      borderSides: ["bottom"],
      border: HAIR,
      borderWidthPt: 0.75,
      paddingSidesPt: { bottom: 8 },
      spacingBeforePt: 8,
      spacingAfterPt: 0,
      hangingIndentMm: 4.94,
      indentMm: 4.94,
      marker: "pointNumber",
    },
    pointNumber: { font: SERIF, weight: "bold", color: GREEN, fontSizePt: 10 },

    // ── Tables of figures ───────────────────────────────────────────────
    /** A plain table: 9pt rows over hairlines. */
    returns: {
      ...text(9),
      borderSides: ["bottom"],
      border: HAIR,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 4, bottom: 4, left: 0, right: 0 },
      valign: "center",
    },
    /** Its head: tracked small labels over a dark line. */
    returnsHead: { ...label(6.5, 0.12, MUTED), border: INK },
    returnStrong: { weight: "bold", color: DEEP },
    returnMuted: { color: MUTED },
    returnGreen: { color: GREEN },
    /** Two lists side by side, each under its own coloured label. */
    leadersGrid: { borderSides: [], paddingPt: 0, valign: "top" },
    leaders: {
      ...text(8.5),
      borderSides: ["bottom"],
      border: HAIR,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 3, bottom: 3, left: 0, right: 0 },
    },
    leadersUp: { ...label(6.5, 0.16, GREEN), spacingBeforePt: 6, spacingAfterPt: 2 },
    leadersDown: { ...label(6.5, 0.16, MUTED), spacingBeforePt: 6, spacingAfterPt: 2 },
    /** The drift table: names, bars against their agreed ranges, and two figures. */
    drift: { ...text(8.5), borderSides: [], paddingPt: 0, valign: "center" },
    driftHead: { ...label(6.5, 0.14, MUTED), valign: "top" },
    driftName: { fontSizePt: 9, weight: "bold" },
    driftNow: { weight: "bold", color: DEEP },
    driftTarget: { color: MUTED },
    driftAfter: { color: GREEN },
    /** The currency bar: five segments, each labelled in a colour it can carry. */
    currency: { borderSides: [], paddingPt: 0, valign: "center", fontSizePt: 7.5, weight: "bold", lineHeight: 1.15 },
    currency0: { fill: DEEP, color: "FFFFFF" },
    currency1: { fill: GREEN, color: "FFFFFF" },
    currency2: { fill: MID, color: "FFFFFF" },
    currency3: { fill: PALE, color: DEEP },
    currency4: { fill: RULE, color: DEEP },
    currencyKey: { ...text(8), spacingBeforePt: 6 },
    key0: { color: DEEP },
    key1: { color: GREEN },
    key2: { color: MID },
    key3: { color: PALE },
    key4: { color: RULE },
    /** The holdings table: a dark head, rows over hairlines, a heavy foot. */
    holdings: {
      ...text(8.5),
      borderSides: ["bottom"],
      border: HAIR,
      borderWidthPt: 0.75,
      // The figures sit right, so their room on the left is free to give:
      // Word sets the tracked heads a hair wider than the browser, and at the
      // canvas's 4pt "WEIGHT" and "100.0%" break in their 12mm column.
      paddingSidesPt: { top: 5, bottom: 5, left: 0.5, right: 4 },
      valign: "center",
    },
    holdingsHead: { ...label(6, 0.14, "FFFFFF"), fill: DEEP, borderSides: [] },
    holding: {},
    holdingAlt: { fill: "F7F9F8" },
    holdingsTotal: {
      weight: "bold",
      borderSides: ["top", "bottom"],
      border: DEEP,
      borderWidthPt: 2,
    },
    // The canvas's 6pt either side; Word and the preview set the longest fund
    // name a hair wider than the browser, and at 6pt it breaks onto a second line.
    holdingName: { paddingSidesPt: { top: 5, bottom: 5, left: 6, right: 3 } },
    holdingNote: { paddingSidesPt: { top: 5, bottom: 5, left: 8, right: 8 } },
    holdingTitle: { weight: "bold" },
    holdingKind: text(6.5, MUTED),
    holdingFigure: {},
    holdingUp: { weight: "bold", color: GREEN },
    holdingDown: { weight: "bold", color: INK },
    holdingYearUp: { color: GREEN },
    holdingYearDown: { color: INK },
    holdingWhy: text(8, INK, 1.35),
    keyValueSmall: { ...figure(16), lineHeight: LEADING, spacingAfterPt: 3 },
    /** A fund panel: a heavy rule over two columns. */
    focus: {
      borderSides: ["top"],
      border: DEEP,
      borderWidthPt: 2,
      paddingSidesPt: { top: 8, bottom: 0, left: 0, right: 0 },
      valign: "top",
    },
    focusName: { ...figure(13, INK), lineHeight: 1.2 },
    focusKind: { ...text(7.5, MUTED), spacingAfterPt: 12 },
    focusValueRow: { borderSides: [], paddingPt: 0, valign: "bottom" },
    focusValue: figure(17),
    focusShare: text(8, MUTED),
    /** The ground a sparkline is drawn on. */
    sparkFrame: { fill: "F7F9F8", borderSides: [], paddingPt: 0, lineHeight: 1 },
    focusSparkNote: { ...text(6.5, MUTED), spacingBeforePt: 6, spacingAfterPt: 6 },
    focusFacts: {
      ...text(8),
      borderSides: ["bottom"],
      border: HAIR,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 2, bottom: 2, left: 0, right: 0 },
    },
    factLabel: { color: MUTED },
    factValue: { weight: "bold" },
    focusLabel: { ...label(6.5, 0.16, MUTED), spacingAfterPt: 3 },
    positions: {
      ...text(8),
      borderSides: ["bottom"],
      border: HAIR,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 2, bottom: 2, left: 0, right: 0 },
      valign: "center",
    },
    positionBar: { paddingPt: 0 },
    focusWhyLabel: { ...label(6.5, 0.16, GREEN), spacingBeforePt: 8.5, spacingAfterPt: 3 },
    focusWhy: text(9, INK, 1.5),
    /** The income chart: twelve columns standing on a dark hairline. */
    incomeChart: {
      borderSides: ["bottom"],
      border: INK,
      borderWidthPt: 0.75,
      paddingPt: 0,
      valign: "bottom",
    },
    incomeFigure: { ...text(7, MUTED), align: "center", spacingAfterPt: 2 },
    incomeFigureNow: { ...text(7, INK), align: "center", spacingAfterPt: 2 },
    incomeAxis: { ...text(7.5, MUTED), align: "center", borderSides: [], paddingSidesPt: { top: 3, bottom: 0, left: 0, right: 0 } },
    /** A label over a hairline, inside a column of a two-column layout. */
    rubricInCell: {
      ...label(7, 0.18),
      borderSides: ["bottom"],
      border: RULE,
      borderWidthPt: 0.75,
      paddingSidesPt: { bottom: 3 },
      spacingBeforePt: 16,
      spacingAfterPt: 7,
    },
    leadersTotal: { weight: "bold", borderSides: [] },
    swatch0: { color: DEEP },
    swatch1: { color: MID },
    swatch2: { color: RULE },
    bridgeTotal: { weight: "bold", color: INK },
    /** The transactions table: tracked heads, rows set a little looser than a list. */
    transactions: {
      ...text(8.5),
      borderSides: ["bottom"],
      border: HAIR,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 5.5, bottom: 5.5, left: 0, right: 6 },
      valign: "center",
    },
    transactionsHead: { ...label(6.5, 0.14, MUTED), border: INK, paddingSidesPt: { top: 3, bottom: 3, left: 0, right: 6 } },
    plain: {},
    disclosures: { ...text(7, MUTED, 1.5), borderSides: [], paddingPt: 0, valign: "top" },
    glossary: { ...text(7.5, INK, 1.45), borderSides: [], paddingSidesPt: { top: 0, bottom: 6, left: 0, right: 0 }, valign: "top" },
    glossaryTerm: { weight: "bold" },
    glossaryMeaning: { color: MUTED },
    costRow: { paddingSidesPt: { top: 5.5, bottom: 5.5, left: 0, right: 6 } },
    costTotal: { weight: "bold", borderSides: [], paddingSidesPt: { top: 3, bottom: 3, left: 0, right: 6 } },
    costFigure: {
      borderSides: ["top"],
      border: RULE,
      borderWidthPt: 0.75,
      paddingSidesPt: { top: 5, bottom: 0, left: 0, right: 0 },
      valign: "top",
    },
    costFigureTotal: {
      borderSides: ["top"],
      border: DEEP,
      borderWidthPt: 1.5,
      paddingSidesPt: { top: 5, bottom: 0, left: 0, right: 0 },
      valign: "top",
    },
    costLabel: { ...label(6.5, 0.14, MUTED), spacingAfterPt: 2 },
    costValue: { ...figure(15), lineHeight: LEADING, spacingAfterPt: 2 },
    costNote: text(7.5, MUTED),
    centred: { valign: "center", borderSides: [], paddingPt: 0 },
    effectFigure: { ...figure(22), lineHeight: LEADING, spacingAfterPt: 6 },
    effectText: { ...text(9, INK, 1.5), spacingAfterPt: 6 },
    effectNote: text(7, MUTED, 1.45),
    /** The calendar years: bars standing on, and hanging from, one baseline. */
    calendar: EMPTY,
    calendarUp: { valign: "bottom" },
    calendarDown: { valign: "top" },
    calendarFigure: { fontSizePt: 8, weight: "bold", color: DEEP, lineHeight: 1.2, align: "center", spacingAfterPt: 3 },
    calendarYear: { ...text(8, MUTED), align: "center" },
    /** The small print under a table. */
    tableNote: { ...text(7, MUTED), spacingBeforePt: 6 },

    // ── Plain text ──────────────────────────────────────────────────────
    small: text(8, MUTED),
    note: text(7, MUTED, 1.45),
    body: text(9, INK, 1.45),
    prose: text(10, INK, 1.45),
    /** The engine's paragraph under a chart. */
    caption: { ...text(9.5, INK, 1.5), spacingBeforePt: 10 },
    /** "Dynamic: <id>" over an engine-written block, in a preview build. */
    dynamicLabel: { ...text(6.5, MUTED), transform: "uppercase", spacingAfterPt: 2 },
    /** A placeholder, in a preview build: the canvas sets them in its quiet ink. */
    standIn: { color: MUTED },

    // ── Bars ────────────────────────────────────────────────────────────
    barDeep: fill(DEEP),
    barGreen: fill(GREEN),
    barMid: fill(MID),
    barPale: fill(PALE),
    barRule: fill(RULE),
    barTrack: fill(TRACK),
    barRange: fill("DBEEE5"),
    barNone: EMPTY,
    barList: { ...text(8.5), borderSides: [], paddingPt: 0, valign: "center" },
    barLabel: {},
    barLabelStrong: { weight: "bold" },
    barValue: { weight: "bold", color: DEEP },
    barValueMuted: { color: MUTED },
    barValuePlain: {},
    barAxis: fill(INK),

    // ── Space ───────────────────────────────────────────────────────────
    // A table has no space before it in Word, so the room above one is an
    // empty paragraph of an exact depth, named for its depth in points.
    ...Object.fromEntries([2, 3, 4, 6, 8, 10, 12, 14, 16, 20, 24, 28, 34, 45].map((pt) => [`gap${pt}`, { heightPt: pt }])),
  },
};
