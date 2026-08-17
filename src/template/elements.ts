/**
 * The elements a document is made of, and the component types that yield them.
 *
 * Each name is both a value and a type: `<Paragraph>` is the element, and
 * `Paragraph` is what a component yielding one is declared as.
 *
 * @module
 */

import type { GraphType, JsonObject } from "../domain/types.ts";
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
  /** What to show wherever the node has not been written yet. */
  placeholder?: string;
}

/** Props for the `<Document>` element, the root every template is built around. */
export interface DocumentProps extends CommonElementProps {
  /** Identifier for the document. */
  id: string;
  /** The document's title. */
  title: string;
  /** Anything to carry alongside the document. */
  metadata?: JsonObject;
  /** The body of the document. */
  children?: Yield;
}

/** Props for the `<Section>` element. */
export interface SectionProps extends CommonElementProps {
  /** The heading printed above the section. */
  title: string;
  /** The nodes the section contains. */
  children?: Yield;
}

/** Props for the `<Paragraph>` element. */
export interface ParagraphProps extends CommonElementProps, PromptProps {
  /** Text children say the same thing; this is for when interpolation reads better. */
  text?: string;
  /** The paragraph's text, as children. */
  children?: Yield<"paragraph">;
}

/** Props for the `<Image>` element. */
export interface ImageProps extends CommonElementProps, PromptProps {
  /** Where the image file lives, relative to the document project. */
  src?: string;
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

/** Props for the `<TableOfContents>` element. */
export interface TableOfContentsProps extends CommonElementProps {
  /** The heading printed above the contents. */
  title?: string;
}

/**
 * A body written once and walked per entry at request time.
 *
 * This is the one structure a build cannot unroll into plain nodes. A branch
 * has two arms and both can be published; a loop has as many arms as the
 * request has entries, and nobody knows that number until a document is
 * written. So the loop is published as a loop.
 */
export interface RepeatProps extends CommonElementProps {
  /** Path to the collection, relative to the request data. */
  over: string;
  /** Name the entry is bound to, readable as `{{ctx.<as>}}`. Defaults to `item`. */
  as?: string;
  /** Name the position is bound to. Defaults to `index`. */
  indexAs?: string;
  /** The body repeated for every entry. */
  children?: Yield;
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
/** A table of contents, built from the sections around it. */
export const TableOfContents: HostElementType<TableOfContentsProps, "tableOfContents"> = host(
  "tableOfContents",
  "TableOfContents",
);
/**
 * A body walked once per entry in a request-time collection.
 *
 * @example
 * ```tsx
 * <Repeat id="charges" over="charges" as="charge">
 *   <Paragraph id="line" text="{{ctx.charge.label}}" />
 * </Repeat>
 * ```
 */
export const Repeat: HostElementType<RepeatProps, "repeat"> = host("repeat", "Repeat");

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
 * A component that yields a `<Graph>`.
 *
 * @typeParam P The props the component takes.
 */
export type Graph<P = Record<never, never>> = Component<P, "graph">;
/**
 * A component that yields a `<TableOfContents>`.
 *
 * @typeParam P The props the component takes.
 */
export type TableOfContents<P = Record<never, never>> = Component<P, "tableOfContents">;
/**
 * A component that yields a `<Repeat>`.
 *
 * @typeParam P The props the component takes.
 */
export type Repeat<P = Record<never, never>> = Component<P, "repeat">;

/**
 * A component free to yield whatever fits, for wrappers and layout pieces.
 *
 * @typeParam P The props the component takes.
 */
export type Nodes<P = Record<never, never>> = Component<P, never>;
