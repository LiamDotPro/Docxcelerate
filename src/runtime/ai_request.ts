/**
 * Asking an {@linkcode AiClient} to write a node.
 *
 * There are two places a dynamic node gets filled in: a build, which has the
 * data in hand, and an engine resolving a published document, which does not
 * have the components any more. They differ in everything except this — how the
 * client is checked, what it is handed, and what it is told when the method it
 * needs is not there. Both had their own copy of that, down to the wording of
 * the errors. A paragraph's, a chart's and an image's answers are checked here too, for the
 * same reason: both places would otherwise pack whatever came back. A client's
 * own failure is wrapped here too, in an {@linkcode AiGenerationError}, so both
 * places name the failed node the same way.
 *
 * Shaping the answer back onto the node stays with the caller, because a build
 * is assembling a node and a resolve is replacing one.
 *
 * @module
 */

import type {
  AiClient,
  AiGraphResult,
  AiImageResult,
  GraphData,
  GraphNode,
  GraphSeries,
  GraphType,
  ImageNode,
  ParagraphNode,
  PromptSpec,
  RuntimeState,
} from "../domain/types.ts";
import { imageSourceOf, rasterTypeOf } from "../render/image_source.ts";

/**
 * A node's prompts as one instruction.
 *
 * @param prompts The prompts, in the order they were set.
 * @returns One line per prompt, kind first.
 */
export function formatPromptText(prompts: PromptSpec[]): string {
  return prompts.map((entry) => `${entry.kind.toUpperCase()}: ${entry.text}`).join("\n");
}

/**
 * An {@linkcode AiClient} method threw or rejected while writing a node.
 *
 * The client's own error is kept whole on `cause`, so a service can tell a
 * rate limit from an outage the way its provider reports it. The message names
 * the node and repeats the cause's message. Nothing from the request is kept:
 * the prompts and the data they were written from stay out of logs.
 *
 * Only a client's failure is wrapped. A node docxcelerate itself refuses (no
 * client, a missing method, an answer that is not text or not a chart) fails
 * with its own error, as before.
 *
 * @example
 * ```ts
 * try {
 *   await resolveDocument(model, createRuntimeState({ data, aiClient }));
 * } catch (err) {
 *   if (err instanceof AiGenerationError) console.error(err.nodeId, err.kind, err.cause);
 *   throw err;
 * }
 * ```
 */
export class AiGenerationError extends Error {
  /** Always `"AiGenerationError"`, so a log line says what failed. */
  override name: string = "AiGenerationError";
  /** The id the client was handed on `request.node.id`, repeat suffix included, such as `note-1`. */
  readonly nodeId: string;
  /** Which kind of node was being written. */
  readonly kind: "paragraph" | "image" | "graph";

  /**
   * Wraps a client's failure.
   *
   * @param nodeId The id of the node the client was writing.
   * @param kind Which kind of node it was.
   * @param cause What the client threw or rejected with, kept as `cause`.
   */
  constructor(nodeId: string, kind: "paragraph" | "image" | "graph", cause: unknown) {
    super(
      `Dynamic ${kind} "${nodeId}" failed to generate: ${cause instanceof Error ? cause.message : String(cause)}`,
      { cause },
    );
    this.nodeId = nodeId;
    this.kind = kind;
  }
}

/**
 * Asks for a dynamic paragraph's prose.
 *
 * @param client The client to ask, if there is one.
 * @param node The paragraph being written.
 * @param prompts Its resolved prompts.
 * @param state Everything reachable while this document is being written.
 * @returns The paragraph text, trimmed and stripped of the characters XML
 * cannot hold.
 * @throws If there is no client, or it answers with something that is not text;
 * an {@linkcode AiGenerationError} if the client throws or rejects.
 */
export async function generateParagraph(
  client: AiClient | undefined,
  node: ParagraphNode,
  prompts: PromptSpec[],
  state: RuntimeState,
): Promise<string> {
  if (!client) {
    throw new Error(`Dynamic paragraph "${node.id}" requires an aiClient.`);
  }

  const asked = inDocumentLanguage(prompts, state);

  let answer: unknown;
  try {
    answer = await client.generateParagraph({
      node,
      prompt: formatPromptText(asked),
      prompts: asked,
      state,
    });
  } catch (cause) {
    throw new AiGenerationError(node.id, "paragraph", cause);
  }

  return writable(node, answer);
}

// What XML 1.0 cannot hold, bar tab, line feed and carriage return: one of
// these in word/document.xml is a file Word refuses to open.
// deno-lint-ignore no-control-regex
const XML_INVALID = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

/**
 * A model's paragraph as text the renderer can pack. An SDK response handed
 * back whole would otherwise pack as "[object Object]", and a trailing newline
 * as a line break nobody wrote.
 */
function writable(node: ParagraphNode, answer: unknown): string {
  if (typeof answer !== "string") {
    throw new Error(
      `Dynamic paragraph "${node.id}" got an answer that is not text: it returned ${kindOf(answer)}.`,
    );
  }

  return answer.replace(XML_INVALID, "").trim();
}

function kindOf(value: unknown): string {
  if (value === null || value === undefined) return String(value);
  if (Array.isArray(value)) return "an array";
  return typeof value === "object" ? "an object" : `a ${typeof value}`;
}

/**
 * Asks for a dynamic image.
 *
 * @param client The client to ask, if there is one.
 * @param node The image being filled.
 * @param prompts Its resolved prompts.
 * @param state Everything reachable while this document is being written.
 * @returns The picture as a `data:` URI on `path`, with only the `alt`, `width`
 * and `height` the client gave.
 * @throws If the client cannot draw, or answers with a picture Word cannot
 * embed; an {@linkcode AiGenerationError} if it throws or rejects.
 */
export async function generateImage(
  client: AiClient | undefined,
  node: ImageNode,
  prompts: PromptSpec[],
  state: RuntimeState,
): Promise<AiImageResult> {
  if (!client?.generateImage) {
    throw new Error(`Dynamic image "${node.id}" requires an aiClient.generateImage method.`);
  }

  // Images are still not told the language: they have no words to write.
  let answer: unknown;
  try {
    answer = await client.generateImage({
      node,
      prompt: formatPromptText(prompts),
      prompts,
      state,
    });
  } catch (cause) {
    throw new AiGenerationError(node.id, "image", cause);
  }

  return drawable(node, answer);
}

/**
 * A client's picture as the `data:` URI the packer embeds, or an error saying
 * why there is none. Anything else packed as a note where the picture should
 * be, with nothing having failed. Only the fields the client gave are kept, so
 * a build spreading this over its node keeps the size the template stated.
 */
function drawable(node: ImageNode, answer: unknown): AiImageResult {
  if (!isRecord(answer)) {
    cannotDraw(node, "it is not an object");
  }

  const { path, bytes, mediaType, alt, width, height } = answer;

  if (path != null && bytes != null) {
    cannotDraw(node, "it has both a path and bytes");
  }

  let uri: unknown = path;

  if (bytes != null) {
    if (!(bytes instanceof Uint8Array)) {
      cannotDraw(node, "bytes is not a Uint8Array");
    }
    if (typeof mediaType !== "string" || mediaType === "") {
      cannotDraw(node, "bytes came without a mediaType");
    }
    uri = `data:${mediaType};base64,${base64Of(bytes)}`;
  }

  if (typeof uri !== "string" || uri === "") {
    cannotDraw(node, "it has neither a data: path nor bytes");
  }

  const source = imageSourceOf(uri);

  // The path is not repeated: a signed URL carries its credentials in it.
  if (source.kind === "url") {
    cannotDraw(node, "its path is not a data: URI; return the picture's bytes instead");
  }

  if (source.kind === "none" || source.bytes.length === 0) {
    cannotDraw(node, "it carries no picture");
  }

  if (!rasterTypeOf(source.mediaType)) {
    cannotDraw(
      node,
      `${describe(source.mediaType)} is not a picture Word embeds; use image/png, image/jpeg, image/gif or image/bmp`,
    );
  }

  return {
    path: uri,
    ...(alt == null ? {} : { alt: alt as string }),
    ...(width == null ? {} : { width: width as number }),
    ...(height == null ? {} : { height: height as number }),
  };
}

/** Base64 in chunks: spreading a multi-megabyte picture into one call overflows the stack. */
function base64Of(bytes: Uint8Array): string {
  let binary = "";
  for (let at = 0; at < bytes.length; at += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(at, at + 0x8000));
  }
  return btoa(binary);
}

function cannotDraw(node: ImageNode, detail: string): never {
  throw new Error(`Dynamic image "${node.id}" got an answer it cannot draw: ${detail}.`);
}

/**
 * Asks for a dynamic chart's data.
 *
 * @param client The client to ask, if there is one.
 * @param node The graph being filled.
 * @param prompts Its resolved prompts.
 * @param state Everything reachable while this document is being written.
 * @returns The answer in the shape the renderer reads: numbers as numbers,
 * categories as text, and only the fields the client actually gave.
 * @throws If the client cannot illustrate, or answers with something that is
 * not a chart; an {@linkcode AiGenerationError} if it throws or rejects.
 */
export async function generateGraph(
  client: AiClient | undefined,
  node: GraphNode,
  prompts: PromptSpec[],
  state: RuntimeState,
): Promise<AiGraphResult> {
  if (!client?.generateGraph) {
    throw new Error(`Dynamic graph "${node.id}" requires an aiClient.generateGraph method.`);
  }

  const asked = inDocumentLanguage(prompts, state);

  let answer: unknown;
  try {
    answer = await client.generateGraph({
      node,
      prompt: formatPromptText(asked),
      prompts: asked,
      state,
    });
  } catch (cause) {
    throw new AiGenerationError(node.id, "graph", cause);
  }

  return plottable(node, answer);
}

/** The chart types a model may name, as a record so a new {@linkcode GraphType} cannot be added without this hearing of it. */
const GRAPH_TYPES: Record<GraphType, true> = {
  bar: true,
  barHorizontal: true,
  line: true,
  area: true,
  pie: true,
  doughnut: true,
  scatter: true,
  radar: true,
  bubble: true,
};

/**
 * A model's chart in the shape the renderer reads, or an error saying where it
 * is not. Model JSON writes numbers as text often enough that a chart packed
 * from it raw is empty in Word with nothing having failed, and it writes `null`
 * where it means "nothing to say", so `null` is read as absent.
 */
function plottable(node: GraphNode, answer: unknown): AiGraphResult {
  if (!isRecord(answer)) {
    refuse(node, "it is not an object");
  }

  const { graphType, data, caption } = answer;
  const result: AiGraphResult = {};

  if (graphType != null) {
    if (typeof graphType !== "string" || !Object.hasOwn(GRAPH_TYPES, graphType)) {
      refuse(node, `graphType ${describe(graphType)} is not a chart type`);
    }
    result.graphType = graphType as GraphType;
  }

  if (data != null) {
    result.data = graphData(node, data);
  }

  if (caption != null) {
    if (typeof caption !== "string") {
      refuse(node, "caption is not text");
    }
    result.caption = caption;
  }

  return result;
}

function graphData(node: GraphNode, data: unknown): GraphData {
  if (!isRecord(data)) {
    refuse(node, "data is not an object");
  }

  const { categories, series, ...others } = data;

  if (!Array.isArray(series)) {
    refuse(node, "data.series is not an array");
  }

  return {
    ...others,
    ...(categories == null ? {} : { categories: labels(node, categories) }),
    series: series.map((entry, index) => {
      const where = `data.series[${index}]`;

      if (!isRecord(entry)) {
        refuse(node, `${where} is not an object`);
      }

      const { values, sizes, ...described } = entry;

      return {
        ...described,
        values: numbers(node, values, `${where}.values`),
        ...(sizes == null ? {} : { sizes: numbers(node, sizes, `${where}.sizes`) }),
      } as unknown as GraphSeries;
    }),
  } as unknown as GraphData;
}

function numbers(node: GraphNode, values: unknown, where: string): (number | null)[] {
  if (!Array.isArray(values)) {
    refuse(node, `${where} is not an array`);
  }

  return values.map((value, index) => {
    if (value === null || (typeof value === "number" && Number.isFinite(value))) {
      return value;
    }

    const parsed = typeof value === "string" && value.trim() !== "" ? Number(value) : NaN;

    if (!Number.isFinite(parsed)) {
      refuse(node, `${where}[${index}] is ${describe(value)}, not a number`);
    }

    return parsed;
  });
}

function labels(node: GraphNode, categories: unknown): string[] {
  if (!Array.isArray(categories)) {
    refuse(node, "data.categories is not an array");
  }

  return categories.map((label, index) => {
    if (typeof label === "string") {
      return label;
    }

    if (typeof label === "number" && Number.isFinite(label)) {
      return String(label);
    }

    return refuse(node, `data.categories[${index}] is ${describe(label)}, not text`);
  });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function describe(value: unknown): string {
  return typeof value === "number" ? String(value) : JSON.stringify(value) ?? String(value);
}

function refuse(node: GraphNode, detail: string): never {
  throw new Error(`Dynamic graph "${node.id}" got an answer it cannot plot: ${detail}.`);
}

/**
 * A document that declared its language gets prose in it without every node
 * having to say so. An image has no words to write, so it is not told.
 */
function inDocumentLanguage(prompts: PromptSpec[], state: RuntimeState): PromptSpec[] {
  if (state.locale === undefined) {
    return prompts;
  }

  const name = new Intl.DisplayNames(["en"], { type: "language" }).of(state.locale);

  return [
    { kind: "system", text: `Write in ${name}, the language of this document.` },
    ...prompts,
  ];
}
