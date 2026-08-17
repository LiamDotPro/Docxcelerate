import type { AiClient, AiGenerationRequest } from "../domain/types.ts";

/**
 * A stand-in AI client, for seeing what a document asks for without asking a
 * model.
 *
 * @module
 */

/**
 * An {@linkcode AiClient} that writes the prompt back rather than answering it.
 *
 * Useful in tests and while wiring a document up: every dynamic paragraph shows
 * exactly what it would have sent.
 */
export class EchoAiClient implements AiClient {
  /**
   * Echoes the node's prompt, prefixed with the node's id.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The prompt, whitespace collapsed.
   */
  generateParagraph(request: AiGenerationRequest): string {
    const compactPrompt = request.prompt.replace(/\s+/g, " ").trim();
    return `[AI draft for ${request.node.id}] ${compactPrompt}`;
  }
}
