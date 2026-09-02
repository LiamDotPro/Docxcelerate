/**
 * Asking an {@linkcode AiClient} to write a node.
 *
 * There are two places a dynamic node gets filled in: a build, which has the
 * data in hand, and an engine resolving a published document, which does not
 * have the components any more. They differ in everything except this — how the
 * client is checked, what it is handed, and what it is told when the method it
 * needs is not there. Both had their own copy of that, down to the wording of
 * the errors.
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
  GraphNode,
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

  return await client.generateParagraph({
    node,
    prompt: formatPromptText(prompts),
    prompts,
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
 * @returns The series to plot, and what to caption them.
 * @throws If the client cannot illustrate.
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

  return await client.generateGraph({
    node,
    prompt: formatPromptText(prompts),
    prompts,
    state,
  });
}
