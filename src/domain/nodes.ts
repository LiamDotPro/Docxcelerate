/**
 * The nodes a built document is made of.
 *
 * Every node is JSON and carries no code, which is what lets a document be
 * built on one machine and rendered on another. A node whose content is not
 * settled yet says so — `mode`, `when`, `prompts` and `derivers` are how it
 * carries the question rather than the answer.
 *
 * @module
 */

import type { Condition, DataReference, DeriverInvocation, PromptSpec } from "./expressions.ts";
import type { DocumentStyle } from "./style.ts";

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
