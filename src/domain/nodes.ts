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
import type { DocumentStyle, TableAlign, TextAlign } from "./style.ts";

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
  | "shape"
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

/**
 * The chart a {@linkcode GraphNode} draws.
 *
 * `bar` stands columns up, which is what a reader means by a bar chart;
 * `barHorizontal` lays them down, for categories whose names are too long to
 * sit under a column. Stacking is not a type of its own — a stacked bar is
 * still a bar, and saying it twice would let a document ask for a stacked pie.
 */
export type GraphType =
  | "bar"
  | "barHorizontal"
  | "line"
  | "area"
  | "pie"
  | "doughnut"
  | "scatter";

/** Where a chart's key sits, or that it has none. */
export type GraphLegend = "none" | "top" | "bottom" | "left" | "right";

/** One plotted run of numbers, and what to call it. */
export interface GraphSeries {
  /**
   * The name printed in the key.
   *
   * A chart with one series often needs no key — its title already says what
   * is plotted — which is why this is optional rather than the series' id.
   */
  label?: string;
  /**
   * The numbers, in the same order as the chart's categories.
   *
   * `null` is a gap rather than a zero: a month with no reading yet is not a
   * month that read nothing, and a line drawn through it would state a figure
   * nobody measured.
   */
  values: (number | null)[];
  /**
   * The colour this series is drawn in, as RGB hex without the `#`.
   *
   * Left out, the series takes its place in the theme's chart palette, which
   * is what keeps a document re-themeable: a colour written here is one the
   * next theme cannot change.
   */
  color?: string;
}

/** The numbers a {@linkcode GraphNode} plots, and what they are counted against. */
export interface GraphData {
  /**
   * What the values are counted against — months, regions, quarters.
   *
   * A `scatter` reads these as its x values instead, so they are numbers
   * written as text there; anything that will not parse counts as its
   * position. Absent, the categories are the positions themselves.
   */
  categories?: string[];
  /** The runs of numbers to plot, in the order they are drawn and keyed. */
  series: GraphSeries[];
}

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

/**
 * A chart, either given its data at build time or handed one by the engine.
 *
 * Packed as a real Word chart rather than a picture of one: the numbers travel
 * with the document, so a reader can select it, restyle it and open its data,
 * and it stays sharp at any zoom and on any printer. That is the whole reason
 * the node carries data instead of an image — a chart flattened to pixels at
 * build time is a chart nobody downstream can do anything with.
 */
export interface GraphNode extends BaseNode {
  /** Discriminator. */
  kind: "graph";
  /** Whether the data is fixed or is to be produced per document. */
  mode: NodeMode;
  /** Which chart to draw. */
  graphType: GraphType;
  /** The series to plot, when they are known at build time. */
  data?: GraphData;
  /** A caption printed beneath the chart. */
  caption?: string;
  /** What to show while the chart has no data yet. */
  placeholder?: string;
  /** How wide the chart is drawn, in points. The text column unless it is said. */
  width?: number;
  /** How deep it is drawn, in points. Seven twelfths of its width unless it is said. */
  height?: number;
  /**
   * Where the key sits.
   *
   * Absent means the renderer decides, which is a key under the plot for
   * anything with more than one series and none for a chart with one — a key
   * naming the only thing on the chart is a line of text saying what the title
   * already said. A pie is the exception: its key names the slices, so it
   * keeps one however few there are.
   */
  legend?: GraphLegend;
  /** Whether the series stack rather than stand beside one another. */
  stacked?: boolean;
  /**
   * How the values are printed, as an OOXML number format — `"#,##0"`,
   * `"0.0%"`, `"£#,##0"`.
   *
   * The same codes Word's own number formatting uses, because they are what
   * the value axis and the data labels are given verbatim.
   */
  numberFormat?: string;
  /** What the category axis counts along. */
  categoryAxisTitle?: string;
  /** What the value axis measures. */
  valueAxisTitle?: string;
  /**
   * Whether each point prints its own figure.
   *
   * Off for most charts — a number on every point is a table drawn badly — and
   * worth turning on for a pie, where the slices are the reading.
   */
  dataLabels?: boolean;
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


/**
 * A drawn box with words on it.
 *
 * The one thing in the model that is a *shape* rather than text: Word draws a
 * real rectangle, and the paragraphs inside it sit on top of the fill rather
 * than beside it. A callout, a status block, a coloured panel a heading stands
 * in — the things a document reaches for when a paragraph with a background is
 * not enough because the box has to be a stated size.
 *
 * That size is the whole difference from a filled table cell, which is the
 * other way to draw a box here and the right one when the box should grow with
 * its text. A shape is a box you have decided the dimensions of; a cell is a
 * box the text decides. Neither is a substitute for the other, and a document
 * that wants the second should use a `<Table>` with one cell in it.
 *
 * What it looks like — the fill, the rule around it, the room inside, how the
 * words sit against its height — is the theme's, through `variant`, exactly as
 * it is for every other node. A shape names what it is; the style says what
 * that looks like.
 */
export interface ShapeNode extends BaseNode {
  /** Discriminator. */
  kind: "shape";
  /**
   * How wide the box is drawn, in points.
   *
   * The text column's full width unless it is said, which is the width a
   * banner or a callout wants and the only width that needs no arithmetic
   * from the document.
   */
  width?: number;
  /**
   * How deep the box is drawn, in points.
   *
   * A shape does not grow to fit its words — that is what makes it a shape —
   * so this is a decision the document or its theme has to make. The theme's
   * `heightPt` is taken when the node says nothing, because a block that
   * declares a depth is already saying exactly this.
   */
  height?: number;
  /** What is drawn on the box: paragraphs, usually, but any node fits. */
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
  | ShapeNode
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
  /**
   * Nodes drawn at the top of left-hand pages, in place of `header`.
   *
   * A document printed on both sides and bound has two kinds of page, not one.
   * The reference belongs at the *outside* edge of each — right on a recto,
   * left on a verso — so it is always the corner a thumb reaches, and a folio
   * that sat in the same place on every sheet would sit in the gutter on half
   * of them.
   *
   * Naming either of these turns on Word's `w:evenAndOddHeaders`, which is a
   * setting for the whole document rather than for one section: from then on
   * `header` and `footer` are what a *right-hand* page shows.
   */
  evenHeader?: DocumentNode[];
  /** Nodes drawn at the foot of left-hand pages, in place of `footer`. */
  evenFooter?: DocumentNode[];
}
