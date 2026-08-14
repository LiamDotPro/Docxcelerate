import type {
  AiClient,
  DocumentModel,
  DocumentNode,
  JsonObject,
  RuntimeState,
} from "./domain/types.ts";
import { InMemoryDataProvider } from "./runtime/data.ts";
import {
  createDeriverRegistry,
  type DeriverDefinitions,
  type DeriverRegistry,
} from "./runtime/derivers.ts";
import type {
  BranchMode,
  ComponentInstance,
  DeriverMode,
  DynamicMode,
  RenderContext,
} from "./template/context.ts";
import type { TemplateElement } from "./template/element.ts";
import type { DocumentProps } from "./template/elements.ts";
import { isPublishValue } from "./template/publish.ts";
import { renderDocumentChildren } from "./template/render.ts";

export interface ComponentRuntimeOptions {
  availableTokens?: number;
  aiClient?: AiClient;
  dynamicMode?: DynamicMode;
  derivers?: DeriverDefinitions | DeriverRegistry;
  deriverMode?: DeriverMode;
  /**
   * Whether a branch is taken now or published for the engine to take.
   *
   * `decide` is the default and is what a preview, a local pack and a live
   * generation all want. `publish` is for the artifact that goes to an engine,
   * where the data is a stand-in and the decision belongs to a request that has
   * not happened yet.
   */
  branchMode?: BranchMode;
  /** How many branches one published document may carry before it is a mistake. */
  branchLimit?: number;
  /** Locale for `useFormat` and placeholder data. */
  locale?: string;
}

export interface DocumentTemplate<TData = unknown> {
  readonly schemaVersion: "docxcelerate.template/v0";
  readonly id: string;
  readonly title: string;
  readonly metadata?: JsonObject;
  readonly element: TemplateElement<"document">;
  /** Present only so the data type survives; never read. */
  readonly __data?: TData;
}

export async function buildDocument<TData>(
  template: DocumentTemplate<TData>,
  data: TData,
  options: ComponentRuntimeOptions = {},
): Promise<DocumentModel> {
  const availableTokens = options.availableTokens ?? 2_000;
  const context = createRenderContext(data, availableTokens, options);
  const props = template.element.props as unknown as DocumentProps;

  return {
    schemaVersion: "docxcelerate.letter/v0",
    id: template.id,
    title: template.title,
    metadata: template.metadata,
    nodes: await renderDocumentChildren(props, context),
  };
}

function createRenderContext<TData>(
  data: TData,
  availableTokens: number,
  options: ComponentRuntimeOptions,
): RenderContext {
  const aiClient = options.aiClient;
  const ctx = dataToObject(data);
  ctx.availableTokens = availableTokens;

  const state: RuntimeState = {
    ctx,
    derived: {},
    dataProvider: new InMemoryDataProvider(ctx),
    aiClient: aiClient ?? missingAiClient,
  };

  return {
    data,
    availableTokens,
    dynamicMode: options.dynamicMode ?? "resolve",
    deriverMode: options.deriverMode ?? "resolve",
    branchMode: options.branchMode ?? "decide",
    branchLimit: options.branchLimit ?? 32,
    locale: options.locale ?? "en-GB",
    derivers: createDeriverRegistry(options.derivers),
    state,
    aiClient,
    shared: new Map<string, unknown>(),
    instances: new Map<string, ComponentInstance>(),
    usedIds: new Map<string, string>(),
    branchesEmitted: 0,
  };
}

/**
 * Spreading the data into `ctx` is what makes `{{data.x}}` resolve. The
 * stand-in used while publishing answers every path and owns no keys, so it is
 * passed through whole rather than copied into an object that would be empty.
 */
function dataToObject<TData>(data: TData): JsonObject {
  if (isPublishValue(data)) {
    return data as unknown as JsonObject;
  }

  if (data && typeof data === "object" && !Array.isArray(data)) {
    return { ...(data as JsonObject) };
  }

  return { value: data };
}

const missingAiClient: AiClient = {
  generateParagraph(): never {
    throw new Error("Dynamic paragraph resolution requires an aiClient.");
  },
};

export type { DocumentNode };
