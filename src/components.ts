import type {
  AiClient,
  DeriverInvocation,
  GraphNode,
  GraphType,
  ImageNode,
  JsonObject,
  LetterDocument,
  LetterNode,
  ParagraphNode,
  PromptKind,
  PromptSpec,
  SectionNode,
  TableOfContentsNode,
} from "./domain/types.ts";
import { InMemoryDataProvider } from "./runtime/data.ts";
import {
  createDeriverRegistry,
  type DeriverDefinitions,
  type DeriverRegistry,
  runDerivers,
} from "./runtime/derivers.ts";
import { renderTemplate } from "./runtime/templates.ts";

type MaybePromise<T> = T | Promise<T>;
type EmptyNode = false | null | undefined;

const nodeComponentMarker = "__docxcelerateNodeComponent";

export interface ComponentRuntimeOptions {
  availableTokens?: number;
  aiClient?: AiClient;
  dynamicMode?: "resolve" | "placeholder";
  derivers?: DeriverDefinitions | DeriverRegistry;
  deriverMode?: "resolve" | "preserve";
}

interface ComponentRuntime {
  availableTokens: number;
  aiClient?: AiClient;
  dynamicMode: "resolve" | "placeholder";
  deriverMode: "resolve" | "preserve";
  derivers: DeriverRegistry;
  state: ReturnType<typeof createRuntimeState>;
}

export interface LetterTemplate<TData> {
  id: string;
  title: string;
  metadata?: JsonObject;
  nodes: NodeComponent<TData>[];
}

export type NodeChildren<TData> = NodeComponent<TData>[];

export type NodeComponent<TData, TNode extends LetterNode = LetterNode> = (
  data: TData,
  availableTokens: number,
) => MaybePromise<Node<TNode, TData> | Node<TNode, TData>[] | EmptyNode>;

export interface TemplateNodeComponent<TData, TNode extends LetterNode = LetterNode> {
  (): TemplateNodeComponent<TData, TNode>;
  (data: TData, availableTokens: number): MaybePromise<
    Node<TNode, TData> | Node<TNode, TData>[] | EmptyNode
  >;
  readonly [nodeComponentMarker]: true;
}

export type Node<TNode extends LetterNode = LetterNode, TData = unknown> = TNode extends SectionNode
  ? SectionDefinition<TData>
  : TNode extends ParagraphNode ? StaticParagraphDefinition | DynamicParagraphDefinition
  : TNode extends ImageNode ? StaticImageDefinition | DynamicImageDefinition
  : TNode extends GraphNode ? StaticGraphDefinition | DynamicGraphDefinition
  : TNode extends TableOfContentsNode ? TableOfContentsDefinition
  :
    | SectionDefinition<TData>
    | StaticParagraphDefinition
    | DynamicParagraphDefinition
    | StaticImageDefinition
    | DynamicImageDefinition
    | StaticGraphDefinition
    | DynamicGraphDefinition
    | TableOfContentsDefinition;

export interface SectionDefinition<TData = unknown> {
  kind: "section";
  id: string;
  title: string;
  nodes: NodeComponent<TData>[];
  derivers?: DeriverInvocation[];
}

export interface StaticParagraphDefinition {
  kind: "paragraph";
  mode: "static";
  id: string;
  render(): MaybePromise<string>;
  derivers?: DeriverInvocation[];
}

interface DynamicPromptDefinition {
  mode: "dynamic";
  id: string;
  derivers?: DeriverInvocation[];
  placeholder?(): MaybePromise<string>;
  generalPrompt(): MaybePromise<string>;
  infoPrompt?(): MaybePromise<string>;
  negativePrompt?(): MaybePromise<string>;
  systemPrompt?(): MaybePromise<string>;
}

export interface DynamicParagraphDefinition extends DynamicPromptDefinition {
  kind: "paragraph";
}

export interface StaticImageDefinition {
  kind: "image";
  mode: "static";
  id: string;
  derivers?: DeriverInvocation[];
  src(): MaybePromise<string>;
  alt?(): MaybePromise<string | undefined>;
  width?(): MaybePromise<number | undefined>;
  height?(): MaybePromise<number | undefined>;
}

export interface DynamicImageDefinition extends DynamicPromptDefinition {
  kind: "image";
}

export interface StaticGraphDefinition {
  kind: "graph";
  mode: "static";
  id: string;
  derivers?: DeriverInvocation[];
  graphType: GraphType;
  data(): MaybePromise<JsonObject>;
  caption?(): MaybePromise<string | undefined>;
}

export interface DynamicGraphDefinition extends DynamicPromptDefinition {
  kind: "graph";
  graphType: GraphType;
}

export interface TableOfContentsDefinition {
  kind: "tableOfContents";
  id: string;
  title?: string;
  derivers?: DeriverInvocation[];
}

export interface DefineLetterOptions<TData> {
  id: string;
  title: string;
  metadata?: JsonObject;
  nodes: NodeChildren<TData>;
}

export function defineLetter<TData>(options: DefineLetterOptions<TData>): LetterTemplate<TData> {
  return letter(
    {
      id: options.id,
      title: options.title,
      metadata: options.metadata,
    },
    options.nodes,
  );
}

export interface LetterOptions {
  id: string;
  title: string;
  metadata?: JsonObject;
}

export function letter<TData>(
  options: LetterOptions,
  nodes: NodeChildren<TData>,
): LetterTemplate<TData> {
  return {
    id: options.id,
    title: options.title,
    metadata: options.metadata,
    nodes,
  };
}

export interface SectionOptions<TData> {
  id: string;
  title: string;
  nodes?: NodeChildren<TData>;
  derivers?: DeriverInvocation[];
}

export function section<TData>(
  options: Omit<SectionOptions<TData>, "nodes">,
  nodes: NodeChildren<TData>,
): NodeComponent<TData, SectionNode>;
export function section<TData>(
  options: SectionOptions<TData>,
): NodeComponent<TData, SectionNode>;
export function section<TData>(
  options: SectionOptions<TData>,
  nodes?: NodeChildren<TData>,
): NodeComponent<TData, SectionNode> {
  const children = nodes ?? options.nodes ?? [];

  return () => ({
    kind: "section",
    id: options.id,
    title: options.title,
    nodes: children,
    derivers: options.derivers,
  });
}

export function isTemplateNodeComponent(value: unknown): value is TemplateNodeComponent<unknown> {
  return Boolean(
    typeof value === "function" &&
      (value as { [nodeComponentMarker]?: boolean })[nodeComponentMarker],
  );
}

function markNodeComponent<TData, TNode extends LetterNode>(
  component: NodeComponent<TData, TNode>,
): TemplateNodeComponent<TData, TNode> {
  Object.defineProperty(component, nodeComponentMarker, {
    value: true,
  });

  return component as TemplateNodeComponent<TData, TNode>;
}

export interface StaticParagraphOptions<TData> {
  id: string;
  derivers?: DeriverInvocation[];
  render(data: TData, availableTokens: number): MaybePromise<string>;
}

export function staticParagraph<TData>(
  options: StaticParagraphOptions<TData>,
): TemplateNodeComponent<TData, ParagraphNode> {
  return markNodeComponent((data, availableTokens) => ({
    kind: "paragraph",
    mode: "static",
    id: options.id,
    derivers: options.derivers,
    render: () => options.render(data, availableTokens),
  }));
}

export interface DynamicParagraphOptions<TData> {
  id: string;
  derivers?: DeriverInvocation[];
  placeholder?(data: TData, availableTokens: number): MaybePromise<string>;
  generalPrompt(data: TData, availableTokens: number): MaybePromise<string>;
  infoPrompt?(data: TData, availableTokens: number): MaybePromise<string>;
  negativePrompt?(data: TData, availableTokens: number): MaybePromise<string>;
  systemPrompt?(data: TData, availableTokens: number): MaybePromise<string>;
}

export function dynamicParagraph<TData>(
  options: DynamicParagraphOptions<TData>,
): TemplateNodeComponent<TData, ParagraphNode> {
  return markNodeComponent((data, availableTokens) => ({
    kind: "paragraph",
    mode: "dynamic",
    id: options.id,
    derivers: options.derivers,
    placeholder: options.placeholder
      ? () => options.placeholder?.(data, availableTokens) ?? ""
      : undefined,
    generalPrompt: () => options.generalPrompt(data, availableTokens),
    infoPrompt: options.infoPrompt
      ? () => options.infoPrompt?.(data, availableTokens) ?? ""
      : undefined,
    negativePrompt: options.negativePrompt
      ? () => options.negativePrompt?.(data, availableTokens) ?? ""
      : undefined,
    systemPrompt: options.systemPrompt
      ? () => options.systemPrompt?.(data, availableTokens) ?? ""
      : undefined,
  }));
}

type StaticValue<TData, TValue> =
  | TValue
  | ((data: TData, availableTokens: number) => MaybePromise<TValue>);

export interface StaticImageOptions<TData> {
  id: string;
  derivers?: DeriverInvocation[];
  src: StaticValue<TData, string>;
  alt?: StaticValue<TData, string | undefined>;
  width?: StaticValue<TData, number | undefined>;
  height?: StaticValue<TData, number | undefined>;
}

export function staticImage<TData>(
  options: StaticImageOptions<TData>,
): TemplateNodeComponent<TData, ImageNode> {
  return markNodeComponent((data, availableTokens) => ({
    kind: "image",
    mode: "static",
    id: options.id,
    derivers: options.derivers,
    src: () => resolveStaticValue(options.src, data, availableTokens),
    alt: options.alt ? () => resolveStaticValue(options.alt, data, availableTokens) : undefined,
    width: options.width
      ? () => resolveStaticValue(options.width, data, availableTokens)
      : undefined,
    height: options.height
      ? () => resolveStaticValue(options.height, data, availableTokens)
      : undefined,
  }));
}

export interface DynamicImageOptions<TData> {
  id: string;
  derivers?: DeriverInvocation[];
  placeholder?(data: TData, availableTokens: number): MaybePromise<string>;
  generalPrompt(data: TData, availableTokens: number): MaybePromise<string>;
  infoPrompt?(data: TData, availableTokens: number): MaybePromise<string>;
  negativePrompt?(data: TData, availableTokens: number): MaybePromise<string>;
  systemPrompt?(data: TData, availableTokens: number): MaybePromise<string>;
}

export function dynamicImage<TData>(
  options: DynamicImageOptions<TData>,
): TemplateNodeComponent<TData, ImageNode> {
  return markNodeComponent((data, availableTokens) => ({
    kind: "image",
    mode: "dynamic",
    id: options.id,
    derivers: options.derivers,
    placeholder: options.placeholder
      ? () => options.placeholder?.(data, availableTokens) ?? ""
      : undefined,
    generalPrompt: () => options.generalPrompt(data, availableTokens),
    infoPrompt: options.infoPrompt
      ? () => options.infoPrompt?.(data, availableTokens) ?? ""
      : undefined,
    negativePrompt: options.negativePrompt
      ? () => options.negativePrompt?.(data, availableTokens) ?? ""
      : undefined,
    systemPrompt: options.systemPrompt
      ? () => options.systemPrompt?.(data, availableTokens) ?? ""
      : undefined,
  }));
}

export interface StaticGraphOptions<TData> {
  id: string;
  derivers?: DeriverInvocation[];
  graphType?: GraphType;
  data: StaticValue<TData, JsonObject>;
  caption?: StaticValue<TData, string | undefined>;
}

export function staticGraph<TData>(
  options: StaticGraphOptions<TData>,
): TemplateNodeComponent<TData, GraphNode> {
  return markNodeComponent((data, availableTokens) => ({
    kind: "graph",
    mode: "static",
    id: options.id,
    derivers: options.derivers,
    graphType: options.graphType ?? "bar",
    data: () => resolveStaticValue(options.data, data, availableTokens),
    caption: options.caption
      ? () => resolveStaticValue(options.caption, data, availableTokens)
      : undefined,
  }));
}

export interface DynamicGraphOptions<TData> {
  id: string;
  derivers?: DeriverInvocation[];
  graphType?: GraphType;
  placeholder?(data: TData, availableTokens: number): MaybePromise<string>;
  generalPrompt(data: TData, availableTokens: number): MaybePromise<string>;
  infoPrompt?(data: TData, availableTokens: number): MaybePromise<string>;
  negativePrompt?(data: TData, availableTokens: number): MaybePromise<string>;
  systemPrompt?(data: TData, availableTokens: number): MaybePromise<string>;
}

export function dynamicGraph<TData>(
  options: DynamicGraphOptions<TData>,
): TemplateNodeComponent<TData, GraphNode> {
  return markNodeComponent((data, availableTokens) => ({
    kind: "graph",
    mode: "dynamic",
    id: options.id,
    derivers: options.derivers,
    graphType: options.graphType ?? "bar",
    placeholder: options.placeholder
      ? () => options.placeholder?.(data, availableTokens) ?? ""
      : undefined,
    generalPrompt: () => options.generalPrompt(data, availableTokens),
    infoPrompt: options.infoPrompt
      ? () => options.infoPrompt?.(data, availableTokens) ?? ""
      : undefined,
    negativePrompt: options.negativePrompt
      ? () => options.negativePrompt?.(data, availableTokens) ?? ""
      : undefined,
    systemPrompt: options.systemPrompt
      ? () => options.systemPrompt?.(data, availableTokens) ?? ""
      : undefined,
  }));
}

export async function buildLetterDocument<TData>(
  template: LetterTemplate<TData>,
  data: TData,
  options: ComponentRuntimeOptions = {},
): Promise<LetterDocument> {
  const aiClient = options.aiClient ?? missingAiClient;
  const runtime: ComponentRuntime = {
    availableTokens: options.availableTokens ?? 2_000,
    aiClient: options.aiClient,
    dynamicMode: options.dynamicMode ?? "resolve",
    deriverMode: options.deriverMode ?? "resolve",
    derivers: createDeriverRegistry(options.derivers),
    state: createRuntimeState(data, aiClient, options.availableTokens ?? 2_000),
  };

  return {
    schemaVersion: "docxcelerate.letter/v0",
    id: template.id,
    title: template.title,
    metadata: template.metadata,
    nodes: await resolveComponents(template.nodes, data, runtime),
  };
}

async function resolveComponents<TData>(
  components: NodeComponent<TData>[],
  data: TData,
  runtime: ComponentRuntime,
): Promise<LetterNode[]> {
  const nodes: LetterNode[] = [];

  for (const component of components) {
    const output = await component(data, runtime.availableTokens);

    for (const definition of flattenDefinitions(output)) {
      nodes.push(await resolveDefinition(definition, data, runtime));
    }
  }

  return nodes;
}

function flattenDefinitions<TData>(
  output: Node<LetterNode, TData> | Node<LetterNode, TData>[] | EmptyNode,
): Node<LetterNode, TData>[] {
  if (!output) {
    return [];
  }

  return Array.isArray(output) ? output.flatMap(flattenDefinitions) : [output];
}

async function resolveDefinition<TData>(
  definition: Node<LetterNode, TData>,
  data: TData,
  runtime: ComponentRuntime,
): Promise<LetterNode> {
  if (definition.kind === "section") {
    await runDefinitionDerivers(definition, runtime);

    return {
      id: definition.id,
      kind: "section",
      title: await renderRequiredRuntimeTemplate(definition.title, runtime),
      derivers: definition.derivers,
      children: await resolveComponents(definition.nodes, data, runtime),
    };
  }

  if (definition.kind === "paragraph") {
    return definition.mode === "static"
      ? await resolveStaticParagraph(definition, runtime)
      : await resolveDynamicParagraph(definition, runtime);
  }

  if (definition.kind === "image") {
    return definition.mode === "static"
      ? await resolveStaticImage(definition, runtime)
      : await resolveDynamicImage(definition, runtime);
  }

  if (definition.kind === "graph") {
    return definition.mode === "static"
      ? await resolveStaticGraph(definition, runtime)
      : await resolveDynamicGraph(definition, runtime);
  }

  await runDefinitionDerivers(definition, runtime);

  return {
    id: definition.id,
    kind: "tableOfContents",
    title: await renderRuntimeTemplate(definition.title, runtime),
    derivers: definition.derivers,
  };
}

async function resolveStaticParagraph(
  definition: StaticParagraphDefinition,
  runtime: ComponentRuntime,
): Promise<ParagraphNode> {
  await runDefinitionDerivers(definition, runtime);

  return {
    id: definition.id,
    kind: "paragraph",
    mode: "static",
    derivers: definition.derivers,
    text: await renderRuntimeTemplate(await definition.render(), runtime),
  };
}

async function resolveDynamicParagraph(
  definition: DynamicParagraphDefinition,
  runtime: ComponentRuntime,
): Promise<ParagraphNode> {
  await runDefinitionDerivers(definition, runtime);

  if (runtime.dynamicMode === "placeholder") {
    return {
      id: definition.id,
      kind: "paragraph",
      mode: "dynamic",
      derivers: definition.derivers,
      text: await resolvePlaceholder(definition, runtime),
    };
  }

  const prompts = await collectPrompts(definition, runtime);
  const node: ParagraphNode = {
    id: definition.id,
    kind: "paragraph",
    mode: "dynamic",
    derivers: definition.derivers,
    prompts,
  };

  if (!runtime.aiClient) {
    throw new Error(`Dynamic paragraph "${definition.id}" requires an aiClient`);
  }

  return {
    ...node,
    text: await runtime.aiClient.generateParagraph({
      node,
      prompt: formatPromptText(prompts),
      prompts,
      state: runtime.state,
    }),
  };
}

async function resolveStaticImage(
  definition: StaticImageDefinition,
  runtime: ComponentRuntime,
): Promise<ImageNode> {
  await runDefinitionDerivers(definition, runtime);

  return {
    id: definition.id,
    kind: "image",
    mode: "static",
    derivers: definition.derivers,
    path: await renderRuntimeTemplate(await definition.src(), runtime),
    alt: await renderRuntimeTemplate(await definition.alt?.(), runtime),
    width: await definition.width?.(),
    height: await definition.height?.(),
  };
}

async function resolveDynamicImage(
  definition: DynamicImageDefinition,
  runtime: ComponentRuntime,
): Promise<ImageNode> {
  await runDefinitionDerivers(definition, runtime);

  if (runtime.dynamicMode === "placeholder") {
    return {
      id: definition.id,
      kind: "image",
      mode: "dynamic",
      derivers: definition.derivers,
      placeholder: await resolvePlaceholder(definition, runtime),
    };
  }

  const prompts = await collectPrompts(definition, runtime);
  const node: ImageNode = {
    id: definition.id,
    kind: "image",
    mode: "dynamic",
    derivers: definition.derivers,
    prompts,
  };

  if (!runtime.aiClient?.generateImage) {
    throw new Error(`Dynamic image "${definition.id}" requires an aiClient.generateImage method`);
  }

  const result = await runtime.aiClient.generateImage({
    node,
    prompt: formatPromptText(prompts),
    prompts,
    state: runtime.state,
  });

  return {
    ...node,
    path: result.path,
    alt: result.alt,
    width: result.width,
    height: result.height,
  };
}

async function resolveStaticGraph(
  definition: StaticGraphDefinition,
  runtime: ComponentRuntime,
): Promise<GraphNode> {
  await runDefinitionDerivers(definition, runtime);

  return {
    id: definition.id,
    kind: "graph",
    mode: "static",
    derivers: definition.derivers,
    graphType: definition.graphType,
    data: await renderJsonObjectRuntimeTemplates(await definition.data(), runtime),
    caption: await renderRuntimeTemplate(await definition.caption?.(), runtime),
  };
}

async function resolveDynamicGraph(
  definition: DynamicGraphDefinition,
  runtime: ComponentRuntime,
): Promise<GraphNode> {
  await runDefinitionDerivers(definition, runtime);

  if (runtime.dynamicMode === "placeholder") {
    return {
      id: definition.id,
      kind: "graph",
      mode: "dynamic",
      derivers: definition.derivers,
      graphType: definition.graphType,
      placeholder: await resolvePlaceholder(definition, runtime),
    };
  }

  const prompts = await collectPrompts(definition, runtime);
  const node: GraphNode = {
    id: definition.id,
    kind: "graph",
    mode: "dynamic",
    derivers: definition.derivers,
    graphType: definition.graphType,
    prompts,
  };

  if (!runtime.aiClient?.generateGraph) {
    throw new Error(`Dynamic graph "${definition.id}" requires an aiClient.generateGraph method`);
  }

  const result = await runtime.aiClient.generateGraph({
    node,
    prompt: formatPromptText(prompts),
    prompts,
    state: runtime.state,
  });

  return {
    ...node,
    graphType: result.graphType ?? definition.graphType,
    data: result.data,
    caption: result.caption,
  };
}

async function resolvePlaceholder(
  definition: DynamicPromptDefinition,
  runtime: ComponentRuntime,
): Promise<string> {
  const placeholder = await definition.placeholder?.();

  if (placeholder && placeholder.trim() !== "") {
    return await renderRequiredRuntimeTemplate(placeholder, runtime);
  }

  return `[Dynamic placeholder: ${definition.id}]`;
}

async function collectPrompts(
  definition: DynamicPromptDefinition,
  runtime: ComponentRuntime,
): Promise<PromptSpec[]> {
  const promptMethods: Array<[PromptKind, (() => MaybePromise<string>) | undefined]> = [
    ["system", definition.systemPrompt],
    ["general", definition.generalPrompt],
    ["info", definition.infoPrompt],
    ["negative", definition.negativePrompt],
  ];
  const prompts: PromptSpec[] = [];

  for (const [kind, method] of promptMethods) {
    if (!method) {
      continue;
    }

    const text = await method();
    if (text.trim() !== "") {
      prompts.push({ kind, text: await renderRequiredRuntimeTemplate(text, runtime) });
    }
  }

  return prompts;
}

function formatPromptText(prompts: PromptSpec[]): string {
  return prompts.map((entry) => `${entry.kind.toUpperCase()}: ${entry.text}`).join("\n");
}

async function runDefinitionDerivers(
  definition: { derivers?: DeriverInvocation[] },
  runtime: ComponentRuntime,
): Promise<void> {
  if (runtime.deriverMode === "preserve") {
    return;
  }

  await runDerivers(definition.derivers, runtime.state, runtime.derivers);
}

async function renderRuntimeTemplate(
  value: string | undefined,
  runtime: ComponentRuntime,
): Promise<string | undefined> {
  if (value === undefined || runtime.deriverMode === "preserve") {
    return value;
  }

  return await renderTemplate(value, runtime.state);
}

async function renderRequiredRuntimeTemplate(
  value: string,
  runtime: ComponentRuntime,
): Promise<string> {
  return await renderRuntimeTemplate(value, runtime) ?? "";
}

async function renderJsonObjectRuntimeTemplates(
  value: JsonObject,
  runtime: ComponentRuntime,
): Promise<JsonObject> {
  return await renderJsonRuntimeTemplates(value, runtime) as JsonObject;
}

async function renderJsonRuntimeTemplates(
  value: unknown,
  runtime: ComponentRuntime,
): Promise<unknown> {
  if (runtime.deriverMode === "preserve") {
    return value;
  }

  if (typeof value === "string") {
    return await renderTemplate(value, runtime.state);
  }

  if (Array.isArray(value)) {
    return await Promise.all(value.map((item) => renderJsonRuntimeTemplates(item, runtime)));
  }

  if (value && typeof value === "object") {
    const entries = await Promise.all(
      Object.entries(value).map(async ([key, nestedValue]) => [
        key,
        await renderJsonRuntimeTemplates(nestedValue, runtime),
      ]),
    );

    return Object.fromEntries(entries);
  }

  return value;
}

function createRuntimeState<TData>(data: TData, aiClient: AiClient, availableTokens: number) {
  const ctx = dataToObject(data);
  ctx.availableTokens = availableTokens;

  return {
    ctx,
    derived: {},
    dataProvider: new InMemoryDataProvider(ctx),
    aiClient,
  };
}

const missingAiClient: AiClient = {
  generateParagraph(): never {
    throw new Error("Dynamic paragraph resolution requires an aiClient.");
  },
};

function resolveStaticValue<TData, TValue>(
  value: StaticValue<TData, TValue>,
  data: TData,
  availableTokens: number,
): MaybePromise<TValue> {
  return typeof value === "function"
    ? (value as (data: TData, availableTokens: number) => MaybePromise<TValue>)(
      data,
      availableTokens,
    )
    : value;
}

function dataToObject<TData>(data: TData): JsonObject {
  if (data && typeof data === "object" && !Array.isArray(data)) {
    return { ...(data as JsonObject) };
  }

  return { value: data };
}
