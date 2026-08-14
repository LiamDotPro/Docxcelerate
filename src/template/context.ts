import type { AiClient, PromptKind, RuntimeState } from "../domain/types.ts";
import type { DeriverRegistry } from "../runtime/derivers.ts";

export type DynamicMode = "resolve" | "placeholder";
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

export interface PromptDraft {
  systemPrompt?: string;
  generalPrompt?: string;
  infoPrompt?: string;
  negativePrompt?: string;
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
