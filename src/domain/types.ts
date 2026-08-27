/**
 * The published document model: the shape a built document takes once the JSX
 * has been evaluated and before an engine turns it into a DOCX.
 *
 * Everything here is data. A `DocumentModel` is JSON — it carries no closures
 * and no component code — which is what lets a document be built on one machine
 * and rendered on another, at a different time, against data that did not exist
 * at build time. The decisions that cannot be made until then survive as
 * {@linkcode Condition}, {@linkcode DeriverInvocation} and {@linkcode RepeatNode}.
 *
 * @module
 */

/** A plain JSON object, used wherever the model carries caller-defined data. */
export type JsonObject = Record<string, unknown>;

/** Discriminator naming which kind of node a {@linkcode DocumentNode} is. */
export type NodeKind =
  | "section"
  | "paragraph"
  | "image"
  | "graph"
  | "table"
  | "tableRow"
  | "tableCell"
  | "tableOfContents"
  | "pageBreak"
  | "pageNumber"
  | "repeat";

/**
 * Whether a node's content is fixed at build time or written per document.
 *
 * `static` content ships in the model as-is. `dynamic` content is left for the
 * engine to fill from the node's prompts and the data it is given.
 */
export type NodeMode = "static" | "dynamic";

/** The chart a {@linkcode GraphNode} draws. */
export type GraphType = "bar" | "line" | "pie";

/**
 * The theme a style came from, by id.
 *
 * A string rather than a union of the shipped themes: the model is data that
 * travels, and a document set in a theme someone wrote themselves should say so
 * rather than claim to be one of ours. The shipped ids are listed by
 * {@linkcode https://docxcelerate.com/themes | the theme catalog} and typed as
 * `ShippedThemeId` in `docxcelerate/themes`.
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
export type PromptKind = "example" | "general" | "info" | "negative" | "system";

/**
 * Which bag of values a {@linkcode DataReference} reads from.
 *
 * `data` is what the caller supplied, `ctx` is what the surrounding repeat or
 * component bound, and `derived` is what a {@linkcode DeriverInvocation} wrote.
 */
export type ReferenceScope = "data" | "ctx" | "derived";

/** A pointer to a value that only exists once a document is being written. */
export interface DataReference {
  /** Which bag of values to read from. */
  scope: ReferenceScope;
  /** A dotted path into that bag, such as `tenant.name`. */
  path: string;
}

/** How the two sides of a `compare` {@linkcode Condition} are measured. */
export type ComparisonOperator = "eq" | "ne" | "gt" | "gte" | "lt" | "lte";

/**
 * A decision the engine makes per document, rather than one the build makes once.
 *
 * `truthy` and `not` came first and are still emitted for the shapes they cover,
 * so an engine that predates the richer forms keeps understanding the common
 * case. The compiler that turns an `if` in a component into one of these picks
 * the narrowest form that fits.
 */
export type Condition =
  | { type: "truthy"; ref: DataReference }
  | { type: "not"; ref: DataReference }
  | {
    type: "compare";
    operator: ComparisonOperator;
    left: ValueExpression;
    right: ValueExpression;
  }
  | { type: "and"; conditions: Condition[] }
  | { type: "or"; conditions: Condition[] }
  | { type: "negate"; condition: Condition };

/**
 * Either a value written into the document at build time or a pointer to one
 * that will not exist until the document is written.
 */
export type ValueExpression =
  | { type: "literal"; value: string | number | boolean }
  | { type: "ref"; ref: DataReference };

/** One instruction attached to a node, for an engine that writes its content. */
export interface PromptSpec {
  /** What the prompt is for. */
  kind: PromptKind;
  /** The prompt itself. */
  text: string;
}

/**
 * A named computation to run before a node is written, and where to put the
 * result.
 *
 * Deriving happens on the engine because the inputs do too. The result lands in
 * the `derived` scope, so anything downstream reads it the same way it reads
 * caller data.
 */
export interface DeriverInvocation {
  /** The key the result is written to under the `derived` scope. */
  output: string;
  /** Which registered deriver to run. */
  name: string;
  /** The arguments to call it with, in order. */
  inputs: ValueExpression[];
}

/** What every node in a document carries, whatever its kind. */
export interface BaseNode {
  /** Identifier, unique within the document. */
  id: string;
  /** Which kind of node this is. */
  kind: NodeKind;
  /** Heading shown for the node, where its kind renders one. */
  title?: string;
  /**
   * Which block style the document's theme should draw this node in.
   *
   * A name, not an appearance: `"band"`, `"panel"`, `"badge"`. What the name
   * looks like is the style's business, which is what keeps a colour out of a
   * component and lets one theme swap for another without touching a node.
   * Unknown names draw as nothing, so a document is never broken by a theme
   * that has not heard of one.
   */
  variant?: string;
  /** Whether the node may be dropped when there is nothing to say. */
  optional?: boolean;
  /** A test that decides, per document, whether the node is included. */
  when?: Condition;
  /** Instructions for an engine writing this node's content. */
  prompts?: PromptSpec[];
  /** Computations to run before the node is written. */
  derivers?: DeriverInvocation[];
}

/** A titled group of nodes. */
export interface SectionNode extends BaseNode {
  /** Discriminator. */
  kind: "section";
  /**
   * Whether a renderer prints the section's title above its children.
   *
   * `false` keeps the title as the section's name — the id it derives, the TOC
   * entry, the address a request targets — without printing it. For a section
   * whose content already says what it is: a charges table whose header row
   * reads "Description", an address block headed by its own label. Absent
   * means printed, which is what a heading is for.
   */
  showTitle?: boolean;
  /** The nodes the section contains, in order. */
  children: DocumentNode[];
}

/** A block of prose, either written at build time or by the engine. */
export interface ParagraphNode extends BaseNode {
  /** Discriminator. */
  kind: "paragraph";
  /** Whether `text` is final or is to be written per document. */
  mode: NodeMode;
  /** The prose, when the paragraph is static or has a fallback. */
  text?: string;
  /**
   * How the paragraph's lines sit in the text column.
   *
   * On the node rather than only in the theme, because alignment is often
   * what the thing *is* rather than how it looks — a date ranged right is
   * ranged right in every theme, the way a money column is. A theme can still
   * say it for a named block, and the node wins when both do, exactly as a
   * cell wins over its column.
   */
  align?: TextAlign;
  /**
   * Pictures set in the line rather than above it.
   *
   * A mark beside a line of credit is one line; given a paragraph of its own
   * it becomes a picture with a caption under it, and a one-line footer bar
   * three lines deep. Each picture records where in `text` it sits, so `text`
   * stays exactly what it was — everything that reads a paragraph's words
   * still reads all of them, in order.
   */
  inlineImages?: InlineImage[];
}

/** A picture set in a paragraph's line, and where along the text it sits. */
export interface InlineImage {
  /** The offset in the paragraph's `text` the picture is placed at. */
  at: number;
  /** The picture itself. */
  image: ImageNode;
}

/** A picture, either supplied at build time or produced by the engine. */
export interface ImageNode extends BaseNode {
  /** Discriminator. */
  kind: "image";
  /** Whether the image is fixed or is to be produced per document. */
  mode: NodeMode;
  /**
   * Where the picture comes from.
   *
   * A `data:` URI carries the bytes in the model, which is the only form that
   * survives being handed to an engine. A path or URL still draws on screen,
   * where a browser can fetch it, but cannot be packed into a Word file.
   */
  path?: string;
  /**
   * A raster to pack in place of an SVG.
   *
   * Word will not embed an SVG without one. Screen renderers ignore this and
   * draw the SVG itself, which is the sharper of the two.
   */
  fallbackPath?: string;
  /** Alternative text describing the image. */
  alt?: string;
  /** Rendered width, in points. */
  width?: number;
  /** Rendered height, in points. */
  height?: number;
  /** What to show while the image has not been produced yet. */
  placeholder?: string;
}

/** A chart, either given its data at build time or handed one by the engine. */
export interface GraphNode extends BaseNode {
  /** Discriminator. */
  kind: "graph";
  /** Whether the data is fixed or is to be produced per document. */
  mode: NodeMode;
  /** Which chart to draw. */
  graphType: GraphType;
  /** The series to plot, when they are known at build time. */
  data?: JsonObject;
  /** A caption printed beneath the chart. */
  caption?: string;
  /** What to show while the chart has no data yet. */
  placeholder?: string;
}

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

/**
 * One column's shape, which every row shares.
 *
 * Widths are declared once, on the table, rather than per cell. A row that set
 * its own would be a row that disagrees with the row above it, and a table
 * whose columns do not line up is not a table.
 */
export interface TableColumn {
  /**
   * Width in millimetres, or `"auto"` to share out what the fixed ones leave.
   *
   * Millimetres rather than points because a page is measured in them: a
   * 26mm money column against a 210mm page is a proportion a reader can check.
   */
  width?: number | "auto";
  /** How this column's cells are aligned. Left unless it is said. */
  align?: TableAlign;
}

/**
 * A grid of cells, with the columns declared once.
 *
 * The rows are children rather than a field, so everything that works on a
 * node works on a row: a loop produces rows per entry, a condition drops one
 * per document, and each carries an id. That is the whole reason a row is a
 * node and not a tuple — an invoice's lines are a `.map()`, and the engine has
 * to be able to walk them without the table being a special case.
 */
export interface TableNode extends BaseNode {
  /** Discriminator. */
  kind: "table";
  /** The columns, left to right. */
  columns: TableColumn[];
  /** The rows, and any loops that produce them. */
  children: DocumentNode[];
}

/** One row of a {@linkcode TableNode}. */
export interface TableRowNode extends BaseNode {
  /** Discriminator. */
  kind: "tableRow";
  /**
   * Whether this row heads the table.
   *
   * A header row is drawn as one and repeats at the top of every page the
   * table runs onto, which is a thing only the renderer can do — a second
   * header written into the body would be a row of text that says the same
   * words in the wrong place.
   */
  header?: boolean;
  /** The cells, left to right. */
  children: DocumentNode[];
}

/** One cell of a {@linkcode TableRowNode}. */
export interface TableCellNode extends BaseNode {
  /** Discriminator. */
  kind: "tableCell";
  /** How many columns this cell runs across. One unless it is said. */
  span?: number;
  /** Alignment, when this cell departs from its column's. */
  align?: TableAlign;
  /** What the cell holds — paragraphs, usually, but any node fits. */
  children: DocumentNode[];
}

/** A table of contents, built from the sections around it. */
export interface TableOfContentsNode extends BaseNode {
  /** Discriminator. */
  kind: "tableOfContents";
}

/**
 * Where one page ends and the next begins.
 *
 * Only ever written where the break is part of what the document *is* — an
 * invoice whose payment details belong on their own page, a contract whose
 * signature block must not be orphaned. Breaking to control where a paragraph
 * happens to land is a job for the margins, not for a node.
 */
export interface PageBreakNode extends BaseNode {
  /** Discriminator. */
  kind: "pageBreak";
}

/** Which page this is, counted while the document is laid out. */
export type PageNumberFormat = "current" | "total" | "currentOfTotal";

/**
 * The page number, filled in by whatever lays the pages out.
 *
 * A build cannot know it: how many pages a document runs to depends on the
 * page size, the font and how much the engine wrote into every dynamic node.
 * So the node says which form it wants and the renderer counts.
 */
export interface PageNumberNode extends BaseNode {
  /** Discriminator. */
  kind: "pageNumber";
  /** Which form to print. Defaults to `currentOfTotal`. */
  format?: PageNumberFormat;
  /** What sits between the two numbers in `currentOfTotal`. Defaults to ` / `. */
  separator?: string;
}

/**
 * A body repeated once per entry in a request-time collection.
 *
 * A build cannot unroll this the way it unrolls a branch: the length of
 * `source` is not known until a document is written. So the loop itself is
 * what gets published, and the engine walks it. Each pass binds the entry under
 * `as` and its position under `indexAs`, both readable through `ctx`, and
 * suffixes child ids with the index so they stay unique across passes.
 */
export interface RepeatNode extends BaseNode {
  /** Discriminator. */
  kind: "repeat";
  /** The collection to walk, resolved per document. */
  source: DataReference;
  /** The `ctx` key each entry is bound to. */
  as: string;
  /** The `ctx` key each entry's zero-based position is bound to. */
  indexAs: string;
  /**
   * A test each entry has to pass to be walked at all.
   *
   * This is what a `.filter()` before the `.map()` becomes. The build cannot
   * apply it — which entries there are belongs to the request — so the test
   * travels with the loop and the engine applies it per entry.
   */
  where?: Condition;
  /** The nodes repeated for every entry, in order. */
  children: DocumentNode[];
}

/** Any node a document can contain. */
export type DocumentNode =
  | SectionNode
  | ParagraphNode
  | ImageNode
  | GraphNode
  | TableNode
  | TableRowNode
  | TableCellNode
  | TableOfContentsNode
  | PageBreakNode
  | PageNumberNode
  | RepeatNode;

/**
 * A built document: the JSON an engine is handed to write one copy from.
 *
 * @example Rendering a built document to a DOCX blob
 * ```ts
 * import { createDocxBlob } from "@docxcelerate/docxcelerate/docx";
 * import type { DocumentModel } from "@docxcelerate/docxcelerate";
 *
 * const model: DocumentModel = {
 *   schemaVersion: "docxcelerate.letter/v0",
 *   id: "welcome",
 *   title: "Welcome",
 *   nodes: [{ id: "hello", kind: "paragraph", mode: "static", text: "Hello." }],
 * };
 *
 * const blob = await createDocxBlob(model);
 * ```
 */
export interface DocumentModel {
  /** The model version, so a reader knows what it is looking at. */
  schemaVersion: "docxcelerate.letter/v0";
  /** Identifier for the document. */
  id: string;
  /** The document's title. */
  title: string;
  /** How the document looks; a renderer default applies when absent. */
  style?: DocumentStyle;
  /** Anything the caller wants to carry alongside the document. */
  metadata?: JsonObject;
  /** The body of the document, in order. */
  nodes: DocumentNode[];
  /**
   * Nodes drawn at the top of every page.
   *
   * Running furniture, not the first thing in the body: it repeats, and it sits
   * outside the text the margins measure. A letterhead that should appear once
   * belongs in `nodes`.
   */
  header?: DocumentNode[];
  /** Nodes drawn at the foot of every page. */
  footer?: DocumentNode[];
  /**
   * Nodes drawn at the top of the first page, in place of `header`.
   *
   * Present only when the document said its first page differs — a letter
   * whose letterhead *is* the top of page one does not want the running strip
   * repeating above it. An empty array means the first page shows nothing
   * where the other pages show `header`.
   */
  firstHeader?: DocumentNode[];
  /** Nodes drawn at the foot of the first page, in place of `footer`. */
  firstFooter?: DocumentNode[];
}

/** Everything a node can reach while a single document is being written. */
export interface RuntimeState {
  /** Values bound by the surrounding repeats and components. */
  ctx: JsonObject;
  /** Values written by derivers that have already run. */
  derived: JsonObject;
  /** Where caller data is read from. */
  dataProvider: DataProvider;
  /** What writes the content of dynamic nodes. */
  aiClient: AiClient;
}

/**
 * Where the `data` scope is read from.
 *
 * An interface rather than a plain object so the data can be fetched lazily —
 * a document that never reads a field never pays to load it.
 */
export interface DataProvider {
  /**
   * Reads one dotted path.
   *
   * @param path The path to read, such as `tenant.name`.
   * @returns The value, or a promise of it. `undefined` when there is none.
   */
  get(path: string): Promise<unknown> | unknown;
}

/** What an {@linkcode AiClient} is told when asked to write a paragraph. */
export interface AiGenerationRequest {
  /** The node being written. */
  node: ParagraphNode;
  /** The node's prompts, joined into one instruction. */
  prompt: string;
  /** The node's prompts, kept apart by kind. */
  prompts: PromptSpec[];
  /** Everything reachable while this document is being written. */
  state: RuntimeState;
}

/** What an {@linkcode AiClient} is told when asked to produce an image. */
export interface AiImageGenerationRequest {
  /** The node being filled. */
  node: ImageNode;
  /** The node's prompts, joined into one instruction. */
  prompt: string;
  /** The node's prompts, kept apart by kind. */
  prompts: PromptSpec[];
  /** Everything reachable while this document is being written. */
  state: RuntimeState;
}

/** The image an {@linkcode AiClient} produced. */
export interface AiImageResult {
  /** Where the image was written. */
  path: string;
  /** Alternative text describing it. */
  alt?: string;
  /** Rendered width, in points. */
  width?: number;
  /** Rendered height, in points. */
  height?: number;
}

/** What an {@linkcode AiClient} is told when asked to produce chart data. */
export interface AiGraphGenerationRequest {
  /** The node being filled. */
  node: GraphNode;
  /** The node's prompts, joined into one instruction. */
  prompt: string;
  /** The node's prompts, kept apart by kind. */
  prompts: PromptSpec[];
  /** Everything reachable while this document is being written. */
  state: RuntimeState;
}

/** The chart an {@linkcode AiClient} produced. Absent fields keep the node's own. */
export interface AiGraphResult {
  /** Which chart to draw. */
  graphType?: GraphType;
  /** The series to plot. */
  data?: JsonObject;
  /** A caption printed beneath the chart. */
  caption?: string;
}

/**
 * What fills in the dynamic nodes of a document.
 *
 * Only paragraphs are required. A client that cannot draw or illustrate simply
 * leaves those methods off, and the resolver falls back to the node's
 * placeholder rather than failing the document.
 */
export interface AiClient {
  /**
   * Writes the prose for a dynamic paragraph.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The paragraph text.
   */
  generateParagraph(request: AiGenerationRequest): Promise<string> | string;
  /**
   * Produces the picture for a dynamic image.
   *
   * @param request The node, its prompts and the state around it.
   * @returns Where the image was written, and how to render it.
   */
  generateImage?(request: AiImageGenerationRequest): Promise<AiImageResult> | AiImageResult;
  /**
   * Produces the data for a dynamic chart.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The series to plot, and optionally the chart type and caption.
   */
  generateGraph?(request: AiGraphGenerationRequest): Promise<AiGraphResult> | AiGraphResult;
}
