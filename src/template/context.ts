/**
 * What a build carries while it renders: the modes it runs in, the state hooks
 * read through, and the per-component bookkeeping behind them.
 *
 * @module
 */

import type { AiClient, PromptKind, RuntimeState } from "../domain/types.ts";
import type { DeriverRegistry } from "../runtime/derivers.ts";

/**
 * Whether a dynamic node is written now or left showing its placeholder.
 *
 * `resolve` calls the AI client. `placeholder` is what a preview does, so
 * building one costs nothing and produces the same document every time.
 */
export type DynamicMode = "resolve" | "placeholder";
/**
 * Whether derivers run during this build or are published for the engine to run.
 *
 * `resolve` runs them against the data at hand. `preserve` keeps the invocation
 * on the node, because the data it needs does not exist yet.
 */
export type DeriverMode = "resolve" | "preserve";

/**
 * Whether a decision belongs to this build or to every document written later.
 *
 * `decide` means data is real and an `if` means what it says. `publish` means
 * the values are stand-ins for a request that has not happened, so a branch is
 * kept rather than taken, and both arms are written out under a condition the
 * engine evaluates per document.
 */
export type BranchMode = "decide" | "publish";

/**
 * The prompts collected for a node while its component renders, before they are
 * written onto the node itself.
 */
export interface PromptDraft {
  /** Standing instructions, such as the voice a document is written in. */
  systemPrompt?: string;
  /** What the node should say. */
  generalPrompt?: string;
  /** Facts to write from. */
  infoPrompt?: string;
  /** What the node must not say. */
  negativePrompt?: string;
  /** What to show wherever the node has not been written yet. */
  placeholder?: string;
}

export const promptPropByKind: Record<PromptKind, keyof PromptDraft> = {
  system: "systemPrompt",
  general: "generalPrompt",
  info: "infoPrompt",
  negative: "negativePrompt",
};

export interface ComponentInstance {
  readonly path: string;
  readonly cells: unknown[];
  cursor: number;
  readonly prompts: PromptDraft;
}

export interface RenderContext {
  readonly data: unknown;
  readonly availableTokens: number;
  readonly dynamicMode: DynamicMode;
  readonly deriverMode: DeriverMode;
  readonly branchMode: BranchMode;
  readonly branchLimit: number;
  readonly locale: string;
  readonly derivers: DeriverRegistry;
  readonly state: RuntimeState;
  readonly aiClient?: AiClient;
  /** Values components leave for the ones rendered after them. */
  readonly shared: Map<string, unknown>;
  readonly instances: Map<string, ComponentInstance>;
  /** Ids already taken, so a collision is reported instead of silently winning. */
  readonly usedIds: Map<string, string>;
  branchesEmitted: number;
}

let currentInstance: ComponentInstance | undefined;
let currentContext: RenderContext | undefined;

export function withInstance<T>(
  context: RenderContext,
  instance: ComponentInstance,
  run: () => T,
): T {
  const previousInstance = currentInstance;
  const previousContext = currentContext;

  instance.cursor = 0;
  currentInstance = instance;
  currentContext = context;

  try {
    return run();
  } finally {
    // Restored synchronously, so anything the component awaits resumes outside
    // the hook context. That is what makes a hook called after an await
    // detectable rather than silently attributed to whoever renders next.
    currentInstance = previousInstance;
    currentContext = previousContext;
  }
}

export function requireInstance(hookName: string): ComponentInstance {
  if (!currentInstance) {
    throw new Error(
      `${hookName} was called outside a component, or after an await inside one. ` +
        "Hooks run in call order, so they must all be reached before the first await " +
        "and never inside a condition or a loop.",
    );
  }

  return currentInstance;
}

export function requireContext(hookName: string): RenderContext {
  if (!currentContext) {
    throw new Error(`${hookName} was called outside a document build.`);
  }

  return currentContext;
}

export function nextCell<T>(instance: ComponentInstance, create: () => T): { value: T } {
  const index = instance.cursor;
  instance.cursor += 1;

  if (index >= instance.cells.length) {
    instance.cells.push({ value: create() });
  }

  return instance.cells[index] as { value: T };
}
