/**
 * How a document looks: the page, the type, and the blocks set on it.
 *
 * A style is data like everything else in the model, so the same document can
 * be packed under one theme today and another tomorrow without the components
 * that wrote it knowing either. `docxcelerate/themes` holds the ready-made
 * ones.
 *
 * @module
 */

export type DocumentStylePreset = string;

/** The paper a document is laid out for. */
export type DocumentPageSize = "A4" | "LETTER";

/** Which way round the page is turned. */
export type DocumentPageOrientation = "portrait" | "landscape";

/** The weight text is set in. */
export type DocumentFontWeight = "regular" | "bold";

/** Whether a block of text is printed as written or shifted to capitals. */
export type DocumentTextTransform = "none" | "uppercase";

/**
 * The colours a theme sets, named by the job each one does rather than by the
 * colour it is.
 *
 * Naming the job is what lets one document be re-themed without touching a
 * node: a rule is drawn in `rule` whichever theme is on, and a theme that wants
 * hairline grey rules says so once here instead of in every node that draws one.
 *
 * Every field is a CSS colour string without the leading `#`, which is the form
 * OOXML wants and the form the web renderer can use unchanged.
 */
export interface DocumentPalette {
  /** Ink for the title and section headings. */
  heading: string;
  /** The one colour the document uses to draw attention. */
  accent: string;
  /** Secondary text: captions, labels, anything set back from the body. */
  muted: string;
  /** Rules, borders and table lines. */
  rule: string;
  /** The paper itself. */
  page: string;
  /**
   * The colours a chart draws its series in, in order.
   *
   * A list rather than a single colour because a chart's colours only mean
   * anything as a set: what matters is that the second series is telling apart
   * from the first, which is a property of the run and not of any one entry.
   * The order is the safety — the hues are stepped so that neighbours stay
   * distinct to a colourblind reader, and cycling or reordering them undoes
   * that.
   *
   * A theme that says nothing gets the shipped default, so a chart is never
   * unstyled. A series that names its own colour overrides whatever stands
   * here, which is the one case a document outranks its theme.
   */
  series?: string[];
}

/** Page margins, in millimetres. */
export interface DocumentPageMargins {
  /** Space above the body text. */
  topMm: number;
  /** Space to the right of the body text. */
  rightMm: number;
  /** Space below the body text. */
  bottomMm: number;
  /** Space to the left of the body text. */
  leftMm: number;
}

/** The page a document is laid out on. */
export interface DocumentPageStyle {
  /** The paper size. */
  size: DocumentPageSize;
  /** Which way round the page is turned. */
  orientation: DocumentPageOrientation;
  /** The margins around the body text. */
  margins: DocumentPageMargins;
  /**
   * How far the running header stands from the top of the *paper*, in
   * millimetres.
   *
   * Measured from the sheet's edge rather than from the margin, because that is
   * what the distance is for: whether a letterhead clears a printer's
   * unprintable edge, whether a running strip collides with a punched hole, how
   * much air stands between the strip and the first line of text. A document
   * that says nothing gets 12.5mm, which is what Word's own default template
   * uses.
   */
  headerMm?: number;
  /** How far the running footer stands from the foot of the paper, in mm. */
  footerMm?: number;
}

/** How body text is set. */
export interface DocumentTypographyStyle {
  /** Font family for body text. */
  bodyFont: string;
  /** Font family for titles and section headings. */
  headingFont: string;
  /** Body text size, in points. */
  bodySizePt: number;
  /** Body line height, as a multiple of the font size. */
  bodyLineHeight: number;
  /** Body text colour, as a CSS colour string. */
  color: string;
}

/** Spacing applied to every paragraph. */
export interface DocumentParagraphStyle {
  /** Space left after each paragraph, in points. */
  spacingAfterPt: number;
}

/** How a heading-like block of text is set apart from the body. */
export interface DocumentTextBlockStyle {
  /** Text size, in points. */
  fontSizePt: number;
  /** The weight the text is set in. */
  weight: DocumentFontWeight;
  /** Space left above the block, in points. */
  spacingBeforePt: number;
  /** Space left below the block, in points. */
  spacingAfterPt: number;
  /** Ink for this block. Falls back to the palette's heading colour. */
  color?: string;
  /** Whether the text is printed as written. Defaults to `none`. */
  transform?: DocumentTextTransform;
  /**
   * Letter spacing in ems, for small capitals that need opening up.
   *
   * Capitals set at a text size are set at the wrong spacing: the letterforms
   * were drawn to sit under lower case, and a label in tracked capitals is
   * what a heading at 7pt has to be to read as one.
   */
  letterSpacingEm?: number;
}

/**
 * Everything about how a document looks, resolved to concrete numbers.
 *
 * A preset is the starting point rather than the whole answer: the style ships
 * with the model, so a renderer needs no access to the preset that produced it.
 */
export interface DocumentStyle {
  /** The theme this style was derived from. */
  preset: DocumentStylePreset;
  /** Page size, orientation and margins. */
  page: DocumentPageStyle;
  /** Fonts, sizes and colour for body text. */
  typography: DocumentTypographyStyle;
  /**
   * The document's colours, by the job each does.
   *
   * Optional because it arrived after the first documents did: a model written
   * before themes existed carries none, and a renderer falls back to the body
   * colour rather than refusing to draw it.
   */
  palette?: DocumentPalette;
  /** Spacing between paragraphs. */
  paragraph: DocumentParagraphStyle;
  /** How the document title is set. */
  title: DocumentTextBlockStyle;
  /**
   * Whether a renderer prints the document's title above the body.
   *
   * On by default, because most documents want their name at the top and
   * should not have to say so. A document that sets its own — an invoice whose
   * letterhead carries the wordmark beside the reference — turns this off, and
   * the title goes on being the document's name for everything that reads the
   * model without being printed twice.
   */
  showTitle?: boolean;
  /** How section headings are set. */
  sectionHeading: DocumentTextBlockStyle;
  /**
   * The block styles a node's `variant` can name.
   *
   * This is the half of the split that says what a name looks like. A component
   * writes `variant="badge"` because the thing *is* a badge; the theme decides
   * that a badge is amber, rounded and set in small capitals. Swapping themes
   * then restyles a document without a node changing.
   */
  blocks?: Record<string, DocumentBlockStyle>;
}

/**
 * How one named block is drawn.
 *
 * Everything is optional and everything omitted is inherited, so a variant that
 * only tints a background says only that.
 */
export interface DocumentBlockStyle {
  /** Background, as a hex string without the `#`. */
  fill?: string;
  /** Text colour, for a block whose fill is dark enough to need one. */
  color?: string;
  /** Border colour. Drawn only when this is set. */
  border?: string;
  /** Border width in points. Defaults to `1` when a border colour is set. */
  borderWidthPt?: number;
  /** Which edges the border is drawn on. Defaults to all four. */
  borderSides?: Array<"top" | "right" | "bottom" | "left">;
  /** Space between the block's edge and its content, in points. */
  paddingPt?: number;
  /**
   * The sides whose padding differs from the rest.
   *
   * A bar that runs to the paper's edge still wants its last words to stop
   * short of it. Naming one side beats inventing a spacer column to hold the
   * gap, which is a column the document does not otherwise have.
   */
  paddingSidesPt?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  };
  /**
   * How deep the block is, in points, when it is a strip rather than type.
   *
   * A rule is a band of colour with no words in it, and its depth is the whole
   * of what it looks like. Without this a theme has to reach that depth by
   * shrinking a font nobody reads until the line collapses around it, which
   * says nothing about what is being drawn.
   */
  heightPt?: number;
  /**
   * Whether the block runs the full width of the page rather than the text.
   *
   * A tinted strip of dates under a letterhead is a band across the sheet, not
   * a box inside the margins — so it reaches past them, on screen by escaping
   * the padding and in Word by indenting negatively.
   */
  bleed?: boolean;
  /** Font size in points, for a block set apart from the body. */
  fontSizePt?: number;
  /** Weight, for a block that carries emphasis. */
  weight?: DocumentFontWeight;
  /** Casing. */
  transform?: DocumentTextTransform;
  /** Letter spacing in ems, for small capitals that need opening up. */
  letterSpacingEm?: number;
  /**
   * The face this block is set in, when it is not the body's.
   *
   * A money column is the reason: proportional digits do not line up under one
   * another, so a figure needs a face whose digits are all one width. Word
   * substitutes a face it does not have exactly as it does for the body font.
   */
  font?: string;
  /**
   * Leading, as a multiple of the font size, when the body's is wrong here.
   *
   * A table row is set tighter than prose and a note under a description
   * tighter still; one leading for a whole document is what makes a row of
   * charges taller than it was drawn.
   */
  lineHeight?: number;
  /**
   * How the block sits against the height of the cell it is in.
   *
   * A band of dates beside a status pill reads as a band only when the two are
   * on the same line as each other, which they are not when a short cell and a
   * tall one both start at the top.
   */
  valign?: "top" | "center" | "bottom";
  /**
   * Space left below the block, in points, when the document's is wrong here.
   *
   * A rule is a strip, not a paragraph of prose: the gap that belongs after a
   * paragraph belongs after prose, and after a hairline it is a hole.
   */
  spacingAfterPt?: number;
  /**
   * The widest the block's lines may run, in millimetres.
   *
   * Prose set across a whole page is prose nobody's eye can track back from;
   * a measure is how a paragraph is kept readable without moving the margin
   * that everything else stands on.
   */
  maxWidthMm?: number;
  /**
   * How this block's lines sit in the width they are given.
   *
   * The theme's half of alignment: a `standfirst` is centred because that is
   * what a standfirst looks like here, and a document that says
   * `variant="standfirst"` should not also have to say `align="center"`. A
   * node that states its own wins.
   */
  align?: TextAlign;
  /**
   * Space left above the block, in points.
   *
   * The counterpart to {@linkcode spacingAfterPt}, and not the same thing.
   * Space above is how a block is set apart from whatever precedes it without
   * the paragraph above having to know that anything follows — which is why
   * Word has both, and why a document built only on space-after ends up with a
   * stray gap at the foot of every page.
   */
  spacingBeforePt?: number;
  /**
   * How far the block is inset from the left margin, in millimetres.
   *
   * Positive numbers only. Reaching *past* the margin is what
   * {@linkcode bleed} is for, and the two would otherwise be two spellings of
   * one thing that disagree about the sign.
   */
  indentMm?: number;
  /** How far the block stops short of the right margin, in millimetres. */
  indentRightMm?: number;
  /**
   * How far the block's first line is indented past the rest, in millimetres.
   *
   * The other way a new paragraph is marked: a book indents rather than
   * leaving a gap, and the two together look like a mistake. Mutually
   * exclusive with {@linkcode hangingIndentMm} — a first line cannot be both
   * pushed in and pulled out — and the hang wins if both are set.
   */
  firstLineIndentMm?: number;
  /**
   * How far the block's first line is pulled back from the rest, in
   * millimetres.
   *
   * What makes a definition's continuation lines clear its term, and what a
   * list marker sits in. Usually the same distance as {@linkcode indentMm}, so
   * the first line starts at the margin and the rest clear it.
   */
  hangingIndentMm?: number;
  /**
   * Whether the block refuses to be the last thing on a page.
   *
   * A heading whose section is overleaf is a heading in the wrong place, and
   * it is not a fault anybody sees in a preview of page one.
   */
  keepWithNext?: boolean;
  /**
   * Whether the block refuses to be split across a page break.
   *
   * For a block that is one thing rather than several — an address, a
   * signature panel — where a break through the middle turns one object into
   * two halves that each look like an accident.
   */
  keepLines?: boolean;
  /**
   * The tab stops this block's text aligns to.
   *
   * A contents line, a signature block, a price beside a description: all of
   * them are one paragraph with a tab in it, and the alternative is a table
   * that draws a grid where a line was wanted.
   */
  tabStopsMm?: TabStop[];
}

/**
 * What a prompt is for, which is how an engine decides where to put it.
 *
 * `general` asks for something, `info` supplies facts to write from, `negative`
 * rules something out, `example` shows the shape a good answer takes, and
 * `system` sets the standing instructions.
 */

/** How a column's cells sit in the width they are given. */
export type TableAlign = "left" | "center" | "right";

/**
 * How a paragraph's lines sit in the width they are given.
 *
 * A superset of {@linkcode TableAlign}: prose can also be justified, which a
 * table column cannot usefully be. Word writes `justify` as `both`, meaning
 * both edges are flush — that translation is the renderer's business, and a
 * document says the word people say.
 */
export type TextAlign = "left" | "center" | "right" | "justify";

/**
 * One tab stop: where it is, what it does to the text that lands on it, and
 * what fills the gap in front of it.
 *
 * The position is measured in millimetres from the left margin, like every
 * other horizontal distance in the model, so a stop at the right margin of an
 * A4 page with 20mm margins is written `170` rather than as a number of twips
 * nobody can check against a ruler.
 */
export interface TabStop {
  /** Where the stop is, in millimetres from the left margin. */
  at: number;
  /** What the text does when it reaches the stop. Left unless it is said. */
  align?: TableAlign | "decimal";
  /**
   * What fills the run-up to the stop.
   *
   * `dot` is a contents line; `none` is the default and is a plain gap.
   */
  leader?: "none" | "dot" | "dash" | "underscore";
}
