/**
 * The one hook for a node an engine writes.
 *
 * Everything a generated node needs is said in a single call, because there is
 * one decision underneath all of it: this node is written per document rather
 * than at build time. Splitting that across several hooks made it look like
 * several decisions, and made the one that is easiest to forget — what a reader
 * sees before anything has been generated — the one you could leave out.
 *
 * So `placeholder` is required. A node with a prompt and no placeholder is not a
 * node with a small gap in it: it is a blank space in every preview, every
 * proof, and every document whose generation failed or was skipped. Making it
 * part of the same call is what stops that from being a thing you remember.
 *
 * @module
 */

import type { JsonObject } from "../domain/types.ts";
import { type PromptDraft, requireInstance } from "./context.ts";

/**
 * What to tell the engine writing a node, and what to show until it has.
 *
 * The names are what each one is for rather than what it is called downstream:
 * a person writing a document is deciding how it should sound and what it may
 * say, not filling in a system prompt.
 */
export interface AiConfig {
  /** What the node should say. */
  ask: string;
  /**
   * What stands in the node's place until something has written it.
   *
   * Required, and deliberately so. A preview is how a document gets proofread,
   * and a document proofread with a blank in it is a document nobody read.
   */
  placeholder: string;
  /** How it should sound — the standing instruction, such as the voice a document is written in. */
  voice?: string;
  /**
   * The facts to write from.
   *
   * Takes the data itself rather than a sentence about it, so the values a
   * component already holds can be handed straight over instead of being
   * formatted into prose first and read back out of it by a model.
   */
  from?: JsonObject | string;
  /** What the node must not say. */
  avoid?: string;
  /**
   * What a good answer looks like, written out.
   *
   * A description of a format is something a model has to interpret; a sample of
   * one is something it can match. So an example is the shortest way to hold
   * still the parts of a node that were never meant to vary — the opening, the
   * order, the length, the register — and leave a model deciding only what is
   * genuinely different from one document to the next.
   *
   * Write it as the finished text, with the per-document parts filled in the way
   * a real one would be. An example full of blanks is a format description again,
   * and it is read as one.
   *
   * Several may be given. One example reads as a template to be followed; two or
   * more read as a pattern to be inferred, which is what to give when the shape
   * legitimately varies by case.
   */
  example?: string | string[];
}

/**
 * Says that this component's node is written per document, and how.
 *
 * Calling this is what makes a node generated; there is no mode to declare. The
 * prompts land on whatever the component yields, so a shared hook can set the
 * house voice on a node it does not own, and a prop written on the element still
 * wins over what a hook set.
 *
 * @param config What to ask for, what to show until it arrives, and the limits.
 * @returns The prompts now set on this component's node.
 * @throws If the placeholder is missing or blank.
 *
 * @example
 * ```tsx
 * export const Summary: Paragraph = () => {
 *   const [invoice] = useState((data: InvoiceData) => data);
 *
 *   useAi({
 *     ask: "Three sentences summarising what this invoice covers.",
 *     placeholder: "A short summary of the work this invoice covers.",
 *     voice: "A delivery lead writing to a finance contact. Plain, no sales tone.",
 *     from: { period: invoice.period, lines: invoice.lines },
 *     avoid: "Do not restate the totals or the payment terms.",
 *     example:
 *       "This invoice covers the June sprint on the payments integration. " +
 *       "The team completed the card capture flow and the refund endpoint. " +
 *       "Testing carried over into July and is billed on the next invoice.",
 *   });
 *
 *   return <Paragraph id="summary" />;
 * };
 * ```
 */
export function useAi(config: AiConfig): PromptDraft {
  const instance = requireInstance("useAi");

  if (typeof config.placeholder !== "string" || config.placeholder.trim() === "") {
    throw new Error(
      "useAi needs a placeholder. It is what a reader sees wherever the node has not been " +
        "written — every preview, every proof, and every document whose generation was " +
        "skipped or failed. Say what should stand there in one line.",
    );
  }

  if (typeof config.ask !== "string" || config.ask.trim() === "") {
    throw new Error(
      "useAi needs an `ask`: what the node should say. A node with a voice and no request " +
        "has nothing to write.",
    );
  }

  Object.assign(instance.prompts, toPromptDraft(config));

  return { ...instance.prompts };
}

/**
 * Turns the words a person writes into the prompts a model is given.
 *
 * The mapping is one-to-one and always has been. What changed is which end
 * names it: `voice` and `avoid` say what they are for, and `systemPrompt` and
 * `negativePrompt` say where they end up. `example` is the one that takes more
 * than one value, and it still lands as one prompt: a node carries a block per
 * kind, so the several are joined here rather than counted everywhere after.
 */
function toPromptDraft(config: AiConfig): PromptDraft {
  const draft: PromptDraft = {
    generalPrompt: config.ask,
    placeholder: config.placeholder,
  };

  if (config.voice !== undefined) {
    draft.systemPrompt = config.voice;
  }

  if (config.avoid !== undefined) {
    draft.negativePrompt = config.avoid;
  }

  if (config.from !== undefined) {
    draft.infoPrompt = typeof config.from === "string" ? config.from : describeFacts(config.from);
  }

  const examples = collectExamples(config.example);

  if (examples !== undefined) {
    draft.examplePrompt = examples;
  }

  return draft;
}

/**
 * Puts one or more examples into the single block a node carries.
 *
 * A lone example travels as itself, because anything wrapped around it is text
 * a model has to decide is not part of the shape it was shown. Several are
 * numbered, since two examples run together read as one long answer, which is
 * the opposite of the point.
 *
 * Blank entries are dropped rather than sent. An empty example is not a weak
 * instruction to a model; it is a demonstration that the right answer is
 * nothing.
 */
function collectExamples(example: string | string[] | undefined): string | undefined {
  if (example === undefined) {
    return undefined;
  }

  const written = (Array.isArray(example) ? example : [example])
    .filter((entry): entry is string => typeof entry === "string")
    .map((entry) => entry.trim())
    .filter((entry) => entry !== "");

  if (written.length === 0) {
    return undefined;
  }

  if (written.length === 1) {
    return written[0];
  }

  return written.map((entry, index) => `Example ${index + 1}:\n${entry}`).join("\n\n");
}

/**
 * Writes structured facts out for a model to read.
 *
 * JSON rather than prose, because the alternative is a component formatting its
 * data into a sentence so that a model can parse the sentence back into data.
 * Undefined entries are dropped rather than written as null, since a fact nobody
 * supplied is not a fact worth stating.
 */
function describeFacts(facts: JsonObject): string {
  return JSON.stringify(facts, (_key, value) => (value === undefined ? undefined : value), 2);
}
