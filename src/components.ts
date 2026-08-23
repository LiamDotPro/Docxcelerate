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
import { renderDocumentChildren, renderDocumentFurniture } from "./template/render.ts";

/**
 * Turning a template into a document: what a build is given, and what it hands
 * back.
 *
 * @module
 */

/** What a build is allowed to do, and what it has to do it with. */
export interface ComponentRuntimeOptions {
  /** The token budget `useAvailableTokens` reports. Defaults to `2000`. */
  availableTokens?: number;
  /** What writes dynamic nodes. Required when `dynamicMode` is `resolve`. */
  aiClient?: AiClient;
  /** Whether dynamic nodes are written now or left as placeholders. */
  dynamicMode?: DynamicMode;
  /** The derivers this build can run. */
  derivers?: DeriverDefinitions | DeriverRegistry;
  /** Whether derivers run now or are published for the engine to run. */
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

/**
 * A `<Document>` tree, ready to build, with the data type it expects attached.
 *
 * @typeParam TData The shape the template reads.
 */
export interface DocumentTemplate<TData = unknown> {
  /** The template version, so a reader knows what it is looking at. */
  readonly schemaVersion: "docxcelerate.template/v0";
  /** The document's identifier, taken from the `<Document>`. */
  readonly id: string;
  /** The document's title, taken from the `<Document>`. */
  readonly title: string;
  /** Anything the `<Document>` carries alongside itself. */
  readonly metadata?: JsonObject;
  /** The root of the element tree. */
  readonly element: TemplateElement<"document">;
  /** Present only so the data type survives; never read. */
  readonly __data?: TData;
}

/**
 * Renders a template into a document model.
 *
 * Components run once each, in document order, so a later one sees what earlier
 * ones left in shared state. What comes out is JSON — how much of it is settled
 * depends on the modes in `options`.
 *
 * @typeParam TData The shape the template reads.
 * @param template The template to build.
 * @param data The data to build it against.
 * @param options What the build is allowed to do.
 * @returns The built document.
 * @throws If a dynamic node needs an AI client and none was given, or two nodes
 * claim the same id.
 */
export async function buildDocument<TData>(
  template: DocumentTemplate<TData>,
  data: TData,
  options: ComponentRuntimeOptions = {},
): Promise<DocumentModel> {
  const availableTokens = options.availableTokens ?? 2_000;
  const context = createRenderContext(data, availableTokens, options);
  const props = template.element.props as unknown as DocumentProps;

  // The body first, so ids are claimed in the order a reader meets them and a
  // collision is reported against the node that actually repeated one.
  const nodes = await renderDocumentChildren(props, context);
  const { header, footer } = await renderDocumentFurniture(props, context);

  return {
    schemaVersion: "docxcelerate.letter/v0",
    id: template.id,
    title: template.title,
    metadata: template.metadata,
    nodes,
    header,
    footer,
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
