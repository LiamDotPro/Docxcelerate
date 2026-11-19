/**
 * Asking an {@linkcode AiClient} to write a node.
 *
 * There are two places a dynamic node gets filled in: a build, which has the
 * data in hand, and an engine resolving a published document, which does not
 * have the components any more. They differ in everything except this — how the
 * client is checked, what it is handed, and what it is told when the method it
 * needs is not there. Both had their own copy of that, down to the wording of
 * the errors. A paragraph's and a chart's answers are checked here too, for the
 * same reason: both places would otherwise pack whatever came back.
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
 * @returns The paragraph text, trimmed and stripped of the characters XML
 * cannot hold.
 * @throws If there is no client, or it answers with something that is not text.
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

  return writable(
    node,
    await client.generateParagraph({
      node,
      prompt: formatPromptText(asked),
      prompts: asked,
      state,
    }),
  );
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
 * The {@linkcode AiGraphResult} a chart client returns, as a JSON Schema a
 * model's structured output can be held to.
 *
 * Pass it as the schema of the provider's structured-output or forced tool
 * call — OpenAI's `response_format` `json_schema` with `strict: true`, an
 * Anthropic tool's `input_schema`, Gemini's `responseJsonSchema` — and return
 * the parsed answer from `generateGraph` as it is. It is written for the
 * strictest of those modes: every field is required, and a field the model
 * has nothing to say for is `null`, which is read as "keep the node's own".
 * The chart types are taken from the same list the answer is checked against,
 * so the two cannot drift. A series' `color` is left out on purpose: a colour
 * a model picks is one the theme can no longer change.
 */
export const aiGraphResultSchema: { readonly type: "object"; readonly [key: string]: unknown } = {
  type: "object",
  additionalProperties: false,
  required: ["graphType", "data", "caption"],
  properties: {
    graphType: {
      description: "The chart to draw, or null to keep the one the document chose.",
      anyOf: [{ type: "string", enum: Object.keys(GRAPH_TYPES) }, { type: "null" }],
    },
    data: {
      type: "object",
      additionalProperties: false,
      required: ["categories", "series"],
      properties: {
        categories: {
          description:
            "What the values are counted against, in order, or null to count by position. For a scatter or bubble chart, the x values written as text.",
          type: ["array", "null"],
          items: { type: "string" },
        },
        series: {
          description: "The runs of numbers to plot, in the order they are drawn.",
          type: "array",
          items: {
            type: "object",
            additionalProperties: false,
            required: ["label", "values", "sizes"],
            properties: {
              label: {
                description: "The name printed in the key, or null for none.",
                type: ["string", "null"],
              },
              values: {
                description: "One number per category, in the same order; null is a gap, not a zero.",
                type: "array",
                items: { type: ["number", "null"] },
              },
              sizes: {
                description: "A bubble chart's point sizes, one per value; null for every other chart.",
                type: ["array", "null"],
                items: { type: ["number", "null"] },
              },
            },
          },
        },
      },
    },
    caption: {
      description: "A caption printed beneath the chart, or null to keep the document's.",
      type: ["string", "null"],
    },
  },
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

      const { values, sizes, label, ...described } = entry;

      if (label != null && typeof label !== "string") {
        refuse(node, `${where}.label is not text`);
      }

      return {
        ...described,
        ...(label == null ? {} : { label }),
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
