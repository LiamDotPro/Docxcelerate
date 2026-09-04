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
  data?: GraphData;
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
