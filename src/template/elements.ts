/**
 * The elements a document is made of, and the component types that yield them.
 *
 * Each name is both a value and a type: `<Paragraph>` is the element, and
 * `Paragraph` is what a component yielding one is declared as.
 *
 * @module
 */

import type {
  GraphType,
  JsonObject,
  PageNumberFormat,
  TableAlign,
  TableColumn,
  TextAlign,
} from "../domain/types.ts";
import {
  type CommonElementProps,
  type Component,
  host,
  type HostElementType,
  type Yield,
} from "./element.ts";

/**
 * Prompt text, written where the node is.
 *
 * Every one of these is also settable with `useSetPrompts`, which is what a
 * shared hook uses to add house style to a node it does not own. Supplying any
 * of them is what makes a node dynamic — the mode is inferred, never declared.
 */
export interface PromptProps {
  /** Standing instructions, such as the voice a document is written in. */
  systemPrompt?: string;
  /** What the node should say. */
  generalPrompt?: string;
  /** Facts to write from. */
  infoPrompt?: string;
  /** What the node must not say. */
  negativePrompt?: string;
  /** What a good answer looks like, shown rather than described. */
  examplePrompt?: string;
  /** What to show wherever the node has not been written yet. */
  placeholder?: string;
}

/** Props for the `<Document>` element, the root every template is built around. */
export interface DocumentProps extends CommonElementProps {
  /**
   * Identifier for the document. Taken from the title when it is left out.
   *
   * Worth writing once the document is being generated somewhere, because it
   * is the name a request asks for. Until then the title says the same thing.
   */
  id?: string;
  /** The document's title. */
  title: string;
  /** Anything to carry alongside the document. */
  metadata?: JsonObject;
  /**
   * Nodes drawn at the top of every page.
   *
   * Running furniture, not the first thing in the body. A letterhead meant to
   * appear once goes in the body; a strip naming the document on every page
   * goes here.
   */
  header?: Yield;
  /** Nodes drawn at the foot of every page — the place a `<PageNumber>` goes. */
  footer?: Yield;
  /**
   * What the top of the *first* page shows, when it differs from `header`.
   *
   * A letter whose letterhead is the top of page one does not want the running
   * strip repeating above it. `false` means the first page shows nothing where
   * the other pages show `header`; nodes mean it shows those instead. Absent
   * means every page gets `header`, which is what running furniture is.
   */
  firstHeader?: Yield | false;
  /** The foot of the first page, when it differs from `footer`. */
  firstFooter?: Yield | false;
  /**
   * What the top of a *left-hand* page shows, when it differs from `header`.
   *
   * For a document printed on both sides and bound. The reference belongs at
   * the outside edge of each page — right on a recto, left on a verso — so it
   * is always the corner a thumb reaches, and a folio in the same place on
   * every sheet sits in the gutter on half of them.
   *
   * Naming either even strip makes `header` and `footer` the right-hand page's.
   */
  evenHeader?: Yield;
  /** The foot of a left-hand page, when it differs from `footer`. */
  evenFooter?: Yield;
  /** The body of the document. */
  children?: Yield;
}

/** Props for the `<Section>` element. */
export interface SectionProps extends CommonElementProps {
  /** The heading printed above the section. */
  title: string;
  /**
   * Whether the title is printed. Defaults to `true`.
   *
   * `false` keeps the title as the section's *name* — the id it derives, its
   * TOC entry, the address a request targets — while printing nothing. For a
   * section whose content already announces itself, like a table whose header
   * row says what the columns are.
   */
  showTitle?: boolean;
  /** The nodes the section contains. */
  children?: Yield;
}

/** Props for the `<Paragraph>` element. */
export interface ParagraphProps extends CommonElementProps, PromptProps {
  /** Text children say the same thing; this is for when interpolation reads better. */
  text?: string;
  /**
   * How the paragraph's lines sit in the text column.
   *
   * Say it here when the alignment is what the paragraph *is* — a date ranged
   * right, a standfirst centred over the piece it opens. Leave it out and let
   * a `variant` carry it when the alignment is what the theme thinks that kind
   * of block looks like. Saying both is allowed, and this wins.
   */
  align?: TextAlign;
  /** The paragraph's text, as children. */
  children?: Yield<"paragraph">;
}

/** Props for the `<Image>` element. */
export interface ImageProps extends CommonElementProps, PromptProps {
  /**
   * Where the picture comes from — a `data:` URI, a path, or a URL.
   *
   * Only a `data:` URI travels: it carries the bytes, so an engine writing the
   * document somewhere else still has the picture.
   */
  src?: string;
  /** A raster to pack in place of an SVG, which Word will not embed alone. */
  fallbackSrc?: string;
  /** Alternative text describing the image. */
  alt?: string;
  /** Rendered width, in points. */
  width?: number;
  /** Rendered height, in points. */
  height?: number;
}

/** Props for the `<Graph>` element. */
export interface GraphProps extends CommonElementProps, PromptProps {
  /** Which chart to draw. */
  graphType?: GraphType;
  /** The series to plot, when they are known at build time. */
  data?: JsonObject;
  /** A caption printed beneath the chart. */
  caption?: string;
}

/** Props for the `<Table>` element. */
export interface TableProps extends CommonElementProps {
  /**
   * The columns, left to right.
   *
   * Declared once here rather than per cell, because every row shares them.
   * A table whose columns do not line up is not a table.
   */
  columns: TableColumn[];
  /** The rows, and any `.map()` producing them. */
  children?: Yield;
}

/** Props for the `<Row>` element. */
export interface RowProps extends CommonElementProps {
  /** Whether this row heads the table, and repeats onto each new page. */
  header?: boolean;
  /** The cells, left to right. */
  children?: Yield;
}

/** Props for the `<Cell>` element. */
export interface CellProps extends CommonElementProps {
  /** How many columns this cell runs across. */
  span?: number;
  /** Alignment, when this cell departs from its column's. */
  align?: TableAlign;
  /**
   * What the cell holds.
   *
   * Text goes straight in, so `<Cell>{line.qty}</Cell>` is the common case and
   * reads as one. Anything needing more than a line — a description above a
   * muted note — writes its paragraphs out instead.
   */
  children?: Yield;
}

/**
 * Props for the `<Shape>` element.
 *
 * @example A callout the theme draws as a filled block
 * ```tsx
 * <Shape id="status" variant="statusBlock" height={44}>
 *   <Paragraph id="status-line">Paid in full — thank you</Paragraph>
 * </Shape>
 * ```
 */
export interface ShapeProps extends CommonElementProps {
  /** How wide the box is drawn, in points. The text column unless it is said. */
  width?: number;
  /** How deep it is drawn, in points. The theme's `heightPt` unless it is said. */
  height?: number;
  /**
   * What is drawn on the box.
   *
   * Text goes straight in, so `<Shape id="x">Paid</Shape>` reads as one line.
   * Anything needing more than a line writes its paragraphs out instead.
   */
  children?: Yield;
}

/** Props for the `<TableOfContents>` element. */
export interface TableOfContentsProps extends CommonElementProps {
  /** The heading printed above the contents. */
  title?: string;
}

/** Props for the `<PageBreak>` element. */
export type PageBreakProps = CommonElementProps;

/** Props for the `<PageNumber>` element. */
export interface PageNumberProps extends CommonElementProps {
  /** Which form to print. Defaults to `currentOfTotal`. */
  format?: PageNumberFormat;
  /** What sits between the two numbers. Defaults to ` / `. */
  separator?: string;
}

// The types below are written out rather than inferred from `host`. Both
// readers need that: a person can see what an element takes without opening
// element.ts, and JSR refuses to publish a public API whose types it would have
// to run the compiler to learn.

/**
 * The root of a template, carrying the document's identity and style.
 *
 * @example
 * ```tsx
 * <Document id="tenancy-renewal" title="Tenancy Renewal">
 *   <Section id="opening" title="Opening" />
 * </Document>
 * ```
 */
export const Document: HostElementType<DocumentProps, "document"> = host(
  "document",
  "Document",
);
/** A titled group of nodes. */
export const Section: HostElementType<SectionProps, "section"> = host("section", "Section");
/**
 * A block of prose. Give it text and it is static; give it a prompt and the
 * engine writes it per document.
 */
export const Paragraph: HostElementType<ParagraphProps, "paragraph"> = host(
  "paragraph",
  "Paragraph",
);
/** A picture, either pointed at a file or left for the engine to produce. */
export const Image: HostElementType<ImageProps, "image"> = host("image", "Image");
/** A chart, either given its data or left for the engine to produce one. */
export const Graph: HostElementType<GraphProps, "graph"> = host("graph", "Graph");
/**
 * A grid of cells, with the columns declared once.
 *
 * @example
 * ```tsx
 * <Table id="lines" columns={[{ width: "auto" }, { width: 26, align: "right" }]}>
 *   <Row header><Cell>Description</Cell><Cell>Amount</Cell></Row>
 *   {state.lines.map((line) => (
 *     <Row><Cell>{line.desc}</Cell><Cell>{currency(line.amount)}</Cell></Row>
 *   ))}
 * </Table>
 * ```
 */
export const Table: HostElementType<TableProps, "table"> = host("table", "Table");
/** One row of a `<Table>`. */
export const Row: HostElementType<RowProps, "tableRow"> = host("tableRow", "Row");
/** One cell of a `<Row>`. */
export const Cell: HostElementType<CellProps, "tableCell"> = host("tableCell", "Cell");

/**
 * A drawn box with words on it.
 *
 * Word draws a real rectangle and the paragraphs sit on top of its fill, which
 * is the difference between this and a paragraph with a background. The box is
 * the size you give it and does not grow with its text — when it should grow,
 * a `<Table>` with one cell in it is the thing to reach for.
 *
 * @example
 * ```tsx
 * <Shape id="status" variant="statusBlock" height={44}>
 *   <Paragraph id="status-line">Paid in full — thank you</Paragraph>
 * </Shape>
 * ```
 */
export const Shape: HostElementType<ShapeProps, "shape"> = host("shape", "Shape");

/**
 * Where one page ends and the next begins.
 *
 * For a break that is part of what the document is — an invoice whose payment
 * details belong on their own page. Nudging a paragraph off the bottom of a
 * page is the margins' job, not a node's.
 */
export const PageBreak: HostElementType<PageBreakProps, "pageBreak"> = host(
  "pageBreak",
  "PageBreak",
);
/** The page number, counted by whatever lays the pages out. */
export const PageNumber: HostElementType<PageNumberProps, "pageNumber"> = host(
  "pageNumber",
  "PageNumber",
);
/** A table of contents, built from the sections around it. */
export const TableOfContents: HostElementType<TableOfContentsProps, "tableOfContents"> = host(
  "tableOfContents",
  "TableOfContents",
);
// The component types below share a name with the element each one yields.
// `const Balance: Paragraph = () => <Paragraph/>` reads as what it is, because
// a value and a type may share a name. Naming the kind is what rejects
// returning a `<Section>` from something declared to be a paragraph.

/**
 * A component that yields a `<Document>`.
 *
 * @typeParam P The props the component takes.
 */
export type Document<P = Record<never, never>> = Component<P, "document">;
/**
 * A component that yields a `<Section>`.
 *
 * @typeParam P The props the component takes.
 */
export type Section<P = Record<never, never>> = Component<P, "section">;
/**
 * A component that yields a `<Paragraph>`.
 *
 * @typeParam P The props the component takes.
 *
 * @example
 * ```tsx
 * export const Greeting: Paragraph = () => <Paragraph id="hello">Hello.</Paragraph>;
 * ```
 */
export type Paragraph<P = Record<never, never>> = Component<P, "paragraph">;
/**
 * A component that yields an `<Image>`.
 *
 * @typeParam P The props the component takes.
 */
export type Image<P = Record<never, never>> = Component<P, "image">;
/**
 * A component that yields a `<Shape>`.
 *
 * @typeParam P The props the component takes.
 */
export type Shape<P = Record<never, never>> = Component<P, "shape">;

/**
 * A component that yields a `<Graph>`.
 *
 * @typeParam P The props the component takes.
 */
export type Graph<P = Record<never, never>> = Component<P, "graph">;
/**
 * A component that yields a `<Table>`.
 *
 * @typeParam P The props the component takes.
 */
export type Table<P = Record<never, never>> = Component<P, "table">;
/**
 * A component that yields a `<Row>`, for a row a document builds in one place.
 *
 * @typeParam P The props the component takes.
 */
export type Row<P = Record<never, never>> = Component<P, "tableRow">;
/**
 * A component that yields a `<Cell>`.
 *
 * @typeParam P The props the component takes.
 */
export type Cell<P = Record<never, never>> = Component<P, "tableCell">;
/**
 * A component that yields a `<PageBreak>`.
 *
 * @typeParam P The props the component takes.
 */
export type PageBreak<P = Record<never, never>> = Component<P, "pageBreak">;
/**
 * A component that yields a `<PageNumber>`.
 *
 * @typeParam P The props the component takes.
 */
export type PageNumber<P = Record<never, never>> = Component<P, "pageNumber">;
/**
 * A component that yields a `<TableOfContents>`.
 *
 * @typeParam P The props the component takes.
 */
export type TableOfContents<P = Record<never, never>> = Component<P, "tableOfContents">;
/**
 * A component free to yield whatever fits, for wrappers and layout pieces.
 *
 * @typeParam P The props the component takes.
 */
export type Nodes<P = Record<never, never>> = Component<P, never>;
