/**
 * What is reachable while a document is being written.
 *
 * This is the only part of the model that is not plain data: an `AiClient` is
 * somebody's implementation and a `DataProvider` reads from somewhere. They
 * are here because a dynamic node cannot be settled without them, and nowhere
 * near the nodes, which stay JSON.
 *
 * @module
 */

import type {
  GraphData,
  GraphNode,
  GraphType,
  ImageNode,
  JsonObject,
  ParagraphNode,
} from "./nodes.ts";
import type { PromptSpec } from "./expressions.ts";

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
  /**
   * The document's locale, as a BCP 47 tag, when it declared one.
   *
   * Here so a dynamic node is written in the document's language, and so a
   * deriver written as a plain `run(inputs, state)` can read it. One written
   * with `deriver()` is called with its inputs alone, and takes the locale as
   * one of them.
   */
  locale?: string;
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
  /** The node being written, its prompts resolved against the request data. */
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
  /** The node being filled, its prompts and alt resolved against the request data. */
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
  /**
   * The picture, as a `data:` URI carrying its bytes. Only that is packed into
   * Word; a path or URL draws on screen but packs as a note naming the image.
   */
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
  /** The node being filled, its prompts, title, alt, caption and axis titles resolved against the request data. */
  node: GraphNode;
  /** The node's prompts, joined into one instruction. */
  prompt: string;
  /** The node's prompts, kept apart by kind. */
  prompts: PromptSpec[];
  /** Everything reachable while this document is being written. */
  state: RuntimeState;
}

/**
 * The chart an {@linkcode AiClient} produced.
 *
 * Absent or `null` fields keep the node's own, so a caption written on the
 * `<Graph>` survives a model that only sends numbers. The answer is read the
 * way model JSON is written: a number written as text, such as `"12"`, is a
 * number, and a numeric category is its text. Anything else that is not a
 * chart fails the document with an error naming the node and the field.
 */
export interface AiGraphResult {
  /** Which chart to draw. */
  graphType?: GraphType;
  /** The series to plot. */
  data?: GraphData;
  /** A caption printed beneath the chart. */
  caption?: string;
}

/**
 * What fills in the dynamic nodes of a document.
 *
 * Only paragraphs are required. A client that cannot draw or illustrate leaves
 * `generateImage` or `generateGraph` off, and a document that never asks for
 * one builds and resolves without it. The first generated image or chart
 * reached without its method fails the build or the resolve with an error
 * naming the node. A placeholder is only what a preview build shows instead of
 * asking at all; it never stands in for a missing method.
 *
 * Every request's `node` arrives with its prompts and the words its template
 * wrote on it already resolved against the request data, on a build and on an
 * engine alike. What the client returns is printed as written: a `{{data.x}}`
 * in a model's answer stays text, so model output cannot pull request data into
 * the document.
 */
export interface AiClient {
  /**
   * Writes the prose for a dynamic paragraph.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The paragraph text, as a string. It is trimmed and stripped of the
   * characters XML cannot hold; anything that is not a string fails the node.
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
