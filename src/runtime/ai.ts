import type { AiClient, AiGenerationRequest } from "../domain/types.ts";

export class EchoAiClient implements AiClient {
  generateParagraph(request: AiGenerationRequest): string {
    const compactPrompt = request.prompt.replace(/\s+/g, " ").trim();
    return `[AI draft for ${request.node.id}] ${compactPrompt}`;
  }
}
