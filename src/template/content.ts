import type { PromptKind, PromptSpec } from "../domain/types.ts";
import { renderTemplate } from "../runtime/templates.ts";
import { type PromptDraft, promptPropByKind, type RenderContext } from "./context.ts";
import { isTemplateElement, type Yield } from "./element.ts";
import type { PromptProps } from "./elements.ts";
import { describe, type Frame } from "./frame.ts";

/**
 * What a node ends up saying.
 *
 * Two jobs that look alike and are not. Text is resolved now: a child that is
 * a string is joined, a template is filled from the data in hand. Prompts are
 * the opposite — they are what a node carries when nobody can write its text
 * yet, and settling them means deciding whether this build answers them or
 * publishes the question for an engine.
 *
 * @module
 */

export function mergePrompts(...drafts: Array<PromptDraft | undefined>): PromptDraft {
  const merged: PromptDraft = {};

  for (const draft of drafts) {
    for (const [key, value] of Object.entries(draft ?? {})) {
      if (value !== undefined) {
        merged[key as keyof PromptDraft] = value as string;
      }
    }
  }

  return merged;
}

/**
 * Decides which prompts a node actually has.
 *
 * A component sets prompts once, before any branch, because hooks run in call
 * order. Its arms need not all want them: one may know exactly what it says.
 * So content settles the question — a node given its own text, source or data
 * is static, and the prompts standing in the air around it were meant for the
 * arm that did not supply any.
 *
 * Saying both on one element is different. That is a single node claiming to be
 * two things, and it is a contradiction rather than a precedence question.
 */
export function settled(
  content: string | undefined,
  props: PromptProps,
  frame: Frame,
  id: string,
): PromptDraft {
  const hasContent = content !== undefined && String(content).trim() !== "";

  if (!hasContent) {
    return mergePrompts(frame.prompts, props);
  }

  if (isDynamic(props)) {
    throw new Error(
      `The node "${id}" at ${describe(frame)} supplies both its content and a prompt to ` +
        "generate it. A node is written or it is generated; supply one of them.",
    );
  }

  return {};
}

export function isDynamic(prompts: PromptProps): boolean {
  return prompts.systemPrompt !== undefined || prompts.generalPrompt !== undefined ||
    prompts.infoPrompt !== undefined || prompts.negativePrompt !== undefined ||
    prompts.examplePrompt !== undefined;
}

export async function promptSpecs(
  prompts: PromptDraft,
  context: RenderContext,
): Promise<PromptSpec[]> {
  // The example reads last because it is the thing the answer is measured
  // against: whatever an engine puts closest to where the writing starts is
  // what the writing ends up shaped like.
  const order: PromptKind[] = ["system", "general", "info", "negative", "example"];
  const specs: PromptSpec[] = [];

  for (const kind of order) {
    const value = prompts[promptPropByKind[kind]];

    if (value !== undefined && value.trim() !== "") {
      specs.push({ kind, text: await requiredText(value, context) });
    }
  }

  return specs;
}

export async function placeholderText(
  prompts: PromptDraft,
  id: string,
  context: RenderContext,
): Promise<string> {
  const placeholder = prompts.placeholder;

  return placeholder && placeholder.trim() !== ""
    ? await requiredText(placeholder, context)
    : `[Dynamic placeholder: ${id}]`;
}

export function formatPromptText(prompts: PromptSpec[]): string {
  return prompts.map((entry) => `${entry.kind.toUpperCase()}: ${entry.text}`).join("\n");
}

export function joinText(children: Yield, frame: Frame): string {
  const parts: string[] = [];

  const walk = (value: Yield): void => {
    if (value === false || value === null || value === undefined) {
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(walk);
      return;
    }

    if (isTemplateElement(value)) {
      throw new Error(
        `A <Paragraph> at ${describe(frame)} was given an element as a child. ` +
          "A paragraph holds text; put elements beside it, not inside it.",
      );
    }

    // Anything else is a value being interpolated. Stringifying rather than
    // testing for `string` is what lets a published build interpolate a
    // stand-in, which is not a string but knows what it is called.
    parts.push(String(value));
  };

  walk(children);

  return parts.join("");
}

export async function text(
  value: string | undefined,
  context: RenderContext,
): Promise<string | undefined> {
  if (value === undefined) {
    return value;
  }

  // A prop can carry a stand-in straight through, so it becomes its own name
  // before anything treats it as text.
  const source = typeof value === "string" ? value : String(value);

  return context.deriverMode === "preserve" ? source : await renderTemplate(source, context.state);
}

export async function requiredText(value: string, context: RenderContext): Promise<string> {
  return await text(value, context) ?? "";
}

export async function jsonText(value: unknown, context: RenderContext): Promise<unknown> {
  if (context.deriverMode === "preserve") {
    return value;
  }

  if (typeof value === "string") {
    return await renderTemplate(value, context.state);
  }

  if (Array.isArray(value)) {
    return await Promise.all(value.map((item) => jsonText(item, context)));
  }

  if (value && typeof value === "object") {
    const entries = await Promise.all(
      Object.entries(value).map(async ([key, nested]) => [key, await jsonText(nested, context)]),
    );

    return Object.fromEntries(entries);
  }

  return value;
}
