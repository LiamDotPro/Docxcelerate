import type {
  AiClient,
  AiGenerationRequest,
  AiGraphGenerationRequest,
  AiGraphResult,
  AiImageGenerationRequest,
  AiImageResult,
  PromptSpec,
} from "../domain/types.ts";

/**
 * Stand-in AI clients, for seeing what a document asks for without asking a
 * model.
 *
 * @module
 */

/**
 * An {@linkcode AiClient} that writes the prompt back rather than answering it.
 *
 * Useful in tests and while wiring a document up: every dynamic paragraph shows
 * exactly what it would have sent, and every dynamic chart and image is filled
 * with a labelled sample, so a document previews whole before any model is
 * plugged in.
 */
export class EchoAiClient implements AiClient {
  /**
   * Echoes the node's prompt, prefixed with the node's id.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The prompt, whitespace collapsed.
   */
  generateParagraph(request: AiGenerationRequest): string {
    return draft(request.node.id, request.prompt);
  }

  /**
   * A one-pixel grey picture, drawn at the size the node asked for.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The sample picture, with the echoed prompt as its alternative text.
   */
  generateImage(request: AiImageGenerationRequest): AiImageResult {
    // The node's own size is handed back because the resolver replaces the
    // node's width and height with the result's, and a sample should not
    // shrink the frame the template reserved.
    return {
      path: SAMPLE_PNG,
      alt: draft(request.node.id, request.prompt),
      width: request.node.width,
      height: request.node.height,
    };
  }

  /**
   * Three sample points, captioned with the echoed prompt.
   *
   * @param request The node, its prompts and the state around it.
   * @returns Sample data that every chart type can draw; the node keeps its own chart type.
   */
  generateGraph(request: AiGraphGenerationRequest): AiGraphResult {
    // Sizes too, because a bubble chart does not draw a point that has none,
    // and every other chart type ignores them.
    return {
      data: {
        categories: ["A", "B", "C"],
        series: [{ label: "Sample", values: [3, 5, 4], sizes: [3, 5, 4] }],
      },
      caption: draft(request.node.id, request.prompt),
    };
  }
}

/** One request a {@linkcode ScriptedAiClient} was handed. */
export interface ScriptedAiRequest {
  /** The node's id as the document wrote it, repeat suffix included, such as `note-0`. */
  id: string;
  /** Which kind of node was asked for. */
  kind: "paragraph" | "image" | "graph";
  /** The node's prompts joined into one instruction: exactly what a model would be sent. */
  prompt: string;
  /** The node's prompts kept apart by kind, including the language line a localised document adds. */
  prompts: PromptSpec[];
}

/**
 * An {@linkcode AiClient} that answers from a script, for tests and offline
 * previews.
 *
 * Answers are keyed by node id: text for a paragraph, an
 * {@linkcode AiGraphResult} for a chart, an {@linkcode AiImageResult} for an
 * image. An id a repeat suffixed, such as `note-0`, falls back to its base id,
 * `note`, when the script has no answer for it exactly. Every request is kept on
 * {@linkcode ScriptedAiClient.requests}, so a test can assert what was asked.
 *
 * @example
 * ```ts
 * const ai = new ScriptedAiClient({ "tutor-note": "A warm note." });
 * await buildDocument(documentTemplate, data, { aiClient: ai });
 * ai.requests[0].prompt; // what the tutor note asked for
 * ```
 */
export class ScriptedAiClient implements AiClient {
  /** Every request this client was handed, in order, including any it had no answer for. */
  readonly requests: ScriptedAiRequest[] = [];
  readonly #answers: Map<string, string | AiGraphResult | AiImageResult>;

  /**
   * Scripts the answers.
   *
   * @param answers What to answer, by node id.
   */
  constructor(answers: Readonly<Record<string, string | AiGraphResult | AiImageResult>>) {
    this.#answers = new Map(Object.entries(answers));
  }

  /**
   * Answers a paragraph from the script.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The scripted text.
   * @throws If the script has no answer for the node, or its answer is not text.
   */
  generateParagraph(request: AiGenerationRequest): string {
    const answer = this.#answer(request);

    if (typeof answer !== "string") {
      throw unusable(request.node.id, "paragraph");
    }

    return answer;
  }

  /**
   * Answers an image from the script.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The scripted image.
   * @throws If the script has no answer for the node, or its answer has no `path`.
   */
  generateImage(request: AiImageGenerationRequest): AiImageResult {
    const answer = this.#answer(request);

    if (typeof answer !== "object" || typeof (answer as AiImageResult).path !== "string") {
      throw unusable(request.node.id, "image");
    }

    return answer as AiImageResult;
  }

  /**
   * Answers a chart from the script.
   *
   * @param request The node, its prompts and the state around it.
   * @returns The scripted chart.
   * @throws If the script has no answer for the node, or its answer is text.
   */
  generateGraph(request: AiGraphGenerationRequest): AiGraphResult {
    const answer = this.#answer(request);

    if (typeof answer !== "object") {
      throw unusable(request.node.id, "graph");
    }

    return answer as AiGraphResult;
  }

  #answer(
    request: AiGenerationRequest | AiImageGenerationRequest | AiGraphGenerationRequest,
  ): string | AiGraphResult | AiImageResult {
    const { id, kind } = request.node;
    // Recorded before answering, so a test can see the request that failed.
    this.requests.push({ id, kind, prompt: request.prompt, prompts: request.prompts });

    // A repeat suffixes ids one `-<index>` per loop, so they are taken off one
    // at a time and an answer for the exact id always wins.
    for (let key = id;;) {
      const answer = this.#answers.get(key);

      if (answer !== undefined) {
        return answer;
      }

      const base = key.replace(/-\d+$/, "");

      if (base === key) {
        break;
      }

      key = base;
    }

    const known = [...this.#answers.keys()].map((key) => `"${key}"`).join(", ");

    throw new Error(
      `ScriptedAiClient has no answer for ${kind} "${id}". Known ids: ${known || "(none)"}.`,
    );
  }
}

function draft(id: string, prompt: string): string {
  return `[AI draft for ${id}] ${prompt.replace(/\s+/g, " ").trim()}`;
}

function unusable(id: string, kind: string): Error {
  return new Error(`ScriptedAiClient's answer for "${id}" is not one a ${kind} can use.`);
}

/** A one-pixel grey PNG: a picture every renderer, Word included, can embed. */
const SAMPLE_PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGO4CQAA2wDaZNW9eAAAAABJRU5ErkJggg==";
