import type { GraphType, JsonObject } from "../domain/types.ts";
import {
  type CommonElementProps,
  type Component,
  host,
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
  systemPrompt?: string;
  generalPrompt?: string;
  infoPrompt?: string;
  negativePrompt?: string;
  placeholder?: string;
}

export interface DocumentProps extends CommonElementProps {
  id: string;
  title: string;
  metadata?: JsonObject;
  children?: Yield;
}

export interface SectionProps extends CommonElementProps {
  title: string;
  children?: Yield;
}

export interface ParagraphProps extends CommonElementProps, PromptProps {
  /** Text children say the same thing; this is for when interpolation reads better. */
  text?: string;
  children?: Yield<"paragraph">;
}

export interface ImageProps extends CommonElementProps, PromptProps {
  src?: string;
  alt?: string;
  width?: number;
  height?: number;
}

export interface GraphProps extends CommonElementProps, PromptProps {
  graphType?: GraphType;
  data?: JsonObject;
  caption?: string;
}

export interface TableOfContentsProps extends CommonElementProps {
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
  children?: Yield;
}

export const Document = host<DocumentProps, "document">("document", "Document");
export const Section = host<SectionProps, "section">("section", "Section");
export const Paragraph = host<ParagraphProps, "paragraph">("paragraph", "Paragraph");
export const Image = host<ImageProps, "image">("image", "Image");
export const Graph = host<GraphProps, "graph">("graph", "Graph");
export const TableOfContents = host<TableOfContentsProps, "tableOfContents">(
  "tableOfContents",
  "TableOfContents",
);
export const Repeat = host<RepeatProps, "repeat">("repeat", "Repeat");

/**
 * The component types, sharing a name with the element each one yields.
 *
 * `const Balance: Paragraph = () => <Paragraph/>` reads as what it is, because
 * a value and a type may share a name. Naming the kind is what rejects
 * returning a `<Section>` from something declared to be a paragraph.
 */
export type Document<P = Record<never, never>> = Component<P, "document">;
export type Section<P = Record<never, never>> = Component<P, "section">;
export type Paragraph<P = Record<never, never>> = Component<P, "paragraph">;
export type Image<P = Record<never, never>> = Component<P, "image">;
export type Graph<P = Record<never, never>> = Component<P, "graph">;
export type TableOfContents<P = Record<never, never>> = Component<P, "tableOfContents">;
export type Repeat<P = Record<never, never>> = Component<P, "repeat">;

/** A component free to yield whatever fits, for wrappers and layout pieces. */
export type Nodes<P = Record<never, never>> = Component<P, never>;
