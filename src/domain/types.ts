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
}

/**
 * What a prompt is for, which is how an engine decides where to put it.
 *
 * `general` asks for something, `info` supplies facts to write from, `negative`
 * rules something out, and `system` sets the standing instructions.
 */
export type PromptKind = "general" | "info" | "negative" | "system";

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
