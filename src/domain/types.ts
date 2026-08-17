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
  | "tableOfContents"
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

/** The named style bundle a document is built from. */
export type DocumentStylePreset = "clean-minimal";

/** The paper a document is laid out for. */
export type DocumentPageSize = "A4" | "LETTER";

/** Which way round the page is turned. */
export type DocumentPageOrientation = "portrait" | "landscape";

/** The weight text is set in. */
export type DocumentFontWeight = "regular" | "bold";

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
}

/**
 * Everything about how a document looks, resolved to concrete numbers.
 *
 * A preset is the starting point rather than the whole answer: the style ships
 * with the model, so a renderer needs no access to the preset that produced it.
 */
export interface DocumentStyle {
  /** The preset this style was derived from. */
  preset: DocumentStylePreset;
  /** Page size, orientation and margins. */
  page: DocumentPageStyle;
  /** Fonts, sizes and colour for body text. */
  typography: DocumentTypographyStyle;
  /** Spacing between paragraphs. */
  paragraph: DocumentParagraphStyle;
  /** How the document title is set. */
  title: DocumentTextBlockStyle;
  /** How section headings are set. */
  sectionHeading: DocumentTextBlockStyle;
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
  /** Where the image file lives, relative to the document project. */
  path?: string;
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

/** A table of contents, built from the sections around it. */
export interface TableOfContentsNode extends BaseNode {
  /** Discriminator. */
  kind: "tableOfContents";
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
  /** The nodes repeated for every entry, in order. */
  children: DocumentNode[];
}

/** Any node a document can contain. */
export type DocumentNode =
  | SectionNode
  | ParagraphNode
  | ImageNode
  | GraphNode
  | TableOfContentsNode
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
