/**
 * Asking an {@linkcode AiClient} to write a node.
 *
 * There are two places a dynamic node gets filled in: a build, which has the
 * data in hand, and an engine resolving a published document, which does not
 * have the components any more. They differ in everything except this — how the
 * client is checked, what it is handed, and what it is told when the method it
 * needs is not there. Both had their own copy of that, down to the wording of
 * the errors. A chart's answer is checked here too, for the same reason: both
 * places would otherwise pack whatever came back.
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
 * Asks for a dynamic paragraph's prose.
 *
 * @param client The client to ask, if there is one.
 * @param node The paragraph being written.
 * @param prompts Its resolved prompts.
 * @param state Everything reachable while this document is being written.
 * @returns The paragraph text.
 * @throws If there is no client.
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

  return await client.generateParagraph({
    node,
    prompt: formatPromptText(asked),
    prompts: asked,
    state,
  });
}

/**
 * Asks for a dynamic image.
 *
 * @param client The client to ask, if there is one.
 * @param node The image being filled.
 * @param prompts Its resolved prompts.
 * @param state Everything reachable while this document is being written.
 * @returns Where the image was written, and how to lay it out.
 * @throws If the client cannot draw.
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

  return await client.generateImage({
    node,
    prompt: formatPromptText(prompts),
    prompts,
    state,
  });
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
 * not a chart.
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

  return plottable(
    node,
    await client.generateGraph({
      node,
      prompt: formatPromptText(asked),
      prompts: asked,
      state,
    }),
  );
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
