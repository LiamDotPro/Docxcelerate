import type {
  Condition,
  DocumentNode,
  GraphNode,
  ImageNode,
  JsonObject,
  ParagraphNode,
  PromptKind,
  PromptSpec,
  RepeatNode,
  SectionNode,
  TableOfContentsNode,
} from "../domain/types.ts";
import { evaluateCondition, invertCondition } from "../runtime/conditions.ts";
import { runDerivers } from "../runtime/derivers.ts";
import { renderTemplate } from "../runtime/templates.ts";
import type { BranchProps } from "./branch.ts";
import {
  type ComponentInstance,
  type PromptDraft,
  promptPropByKind,
  type RenderContext,
  withInstance,
} from "./context.ts";
import {
  type CommonElementProps,
  hostKindOf,
  isTemplateElement,
  type TemplateElement,
  type Yield,
} from "./element.ts";
import type {
  DocumentProps,
  GraphProps,
  ImageProps,
  ParagraphProps,
  PromptProps,
  RepeatProps,
  SectionProps,
  TableOfContentsProps,
} from "./elements.ts";

interface Frame {
  /** Where this sits in the tree. Identity for hooks, and a fallback id. */
  readonly path: string;
  /** Conditions gathered from branches above, carried onto published nodes. */
  readonly when?: Condition;
  /** Prompts set by the component that yielded this element. */
  readonly prompts?: PromptDraft;
  /**
   * Which pass of a loop this is.
   *
   * A body written once and walked many times names its nodes once, so the pass
   * has to distinguish them. The engine suffixes published loops the same way,
   * which is what keeps a previewed id and a written one the same string.
   */
  readonly idSuffix?: string;
}

export async function renderDocumentChildren(
  props: DocumentProps,
  context: RenderContext,
): Promise<DocumentNode[]> {
  return await renderYield(props.children, context, { path: "" });
}

export async function renderYield(
  value: Yield,
  context: RenderContext,
  frame: Frame,
): Promise<DocumentNode[]> {
  const resolved = await value;

  if (resolved === false || resolved === null || resolved === undefined || resolved === "") {
    return [];
  }

  if (Array.isArray(resolved)) {
    const nodes: DocumentNode[] = [];

    for (const [index, child] of resolved.entries()) {
      nodes.push(...await renderYield(child, context, childFrame(frame, index)));
    }

    return nodes;
  }

  if (typeof resolved === "string" || typeof resolved === "number") {
    throw new Error(
      `Text appeared at ${describe(frame)} where a document element belongs. ` +
        "Text only lives inside a <Paragraph>.",
    );
  }

  if (!isTemplateElement(resolved)) {
    throw new Error(
      `${describe(frame)} produced ${typeof resolved}, which is not a document element.`,
    );
  }

  return await renderElement(resolved, context, frame);
}

async function renderElement(
  element: TemplateElement,
  context: RenderContext,
  frame: Frame,
): Promise<DocumentNode[]> {
  if (element.kind === "component") {
    return await renderComponent(element, context, frame);
  }

  if (element.kind === "fragment") {
    return await renderYield(element.props.children as Yield, context, frame);
  }

  if (element.kind === "branch") {
    return await renderBranch(element.props as unknown as BranchProps, context, frame);
  }

  if (element.kind === "document") {
    throw new Error(
      "<Document> can only be the root of a template, not a child of another element.",
    );
  }

  return [await renderHost(element, context, frame)].flat();
}

/**
 * Runs a component.
 *
 * The instance carries the hook cells and the prompt draft, and is keyed by
 * where the component sits, so the same position keeps the same state for the
 * whole build. Whatever prompts the component set are handed to the element it
 * yields — which is why `useSetPrompts` needs no return value at the call site.
 */
async function renderComponent(
  element: TemplateElement,
  context: RenderContext,
  frame: Frame,
): Promise<DocumentNode[]> {
  const component = element.type as (props: Record<string, unknown>) => Yield | Promise<Yield>;
  const name = component.name || "component";
  const path = `${frame.path}/${name}`;
  const instance = instanceAt(context, path);
  const yielded = withInstance(context, instance, () => component(element.props));

  return await renderYield(await yielded, context, {
    path,
    when: frame.when,
    idSuffix: frame.idSuffix,
    prompts: mergePrompts(frame.prompts, instance.prompts),
  });
}

/**
 * Expands a compiled `if`.
 *
 * With real data the condition is evaluated and one arm is kept, which is what
 * the `if` in the source meant. Publishing keeps both arms, each carrying the
 * condition that selects it, so the decision travels to the engine and is made
 * per document instead of once for everybody.
 */
async function renderBranch(
  props: BranchProps,
  context: RenderContext,
  frame: Frame,
): Promise<DocumentNode[]> {
  if (context.branchMode === "decide") {
    const taken = await evaluateCondition(props.condition, context.state);
    const arm = taken ? props.whenTrue : props.whenFalse;

    return arm ? await renderYield(arm(), context, childFrame(frame, taken ? 0 : 1)) : [];
  }

  context.branchesEmitted += 1;

  if (context.branchesEmitted > context.branchLimit) {
    throw new Error(
      `Publishing ${describe(frame)} passed ${context.branchLimit} branches. ` +
        "Every branch publishes both of its arms, so nesting them multiplies what a document " +
        "carries. Lift the decision into a deriver that yields one value, or raise branchLimit " +
        "if the document really is this conditional.",
    );
  }

  const nodes: DocumentNode[] = [];

  if (props.whenTrue) {
    nodes.push(...await renderYield(props.whenTrue(), context, {
      ...childFrame(frame, 0),
      when: allOf(frame.when, props.condition),
    }));
  }

  if (props.whenFalse) {
    nodes.push(...await renderYield(props.whenFalse(), context, {
      ...childFrame(frame, 1),
      when: allOf(frame.when, invertCondition(props.condition)),
    }));
  }

  return nodes;
}

async function renderHost(
  element: TemplateElement,
  context: RenderContext,
  frame: Frame,
): Promise<DocumentNode | DocumentNode[]> {
  const kind = hostKindOf(element.type);
  const props = element.props as CommonElementProps;
  const id = claimId(context, props.id, frame, kind ?? "node");
  const when = allOf(frame.when, props.when);

  await runNodeDerivers(props, context);

  return withDerivers(
    await renderHostKind(kind, element, context, frame, id, when),
    props.derivers,
  );
}

async function renderHostKind(
  kind: string | undefined,
  element: TemplateElement,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<DocumentNode | DocumentNode[]> {
  if (kind === "section") {
    return await renderSection(element.props as unknown as SectionProps, context, frame, id, when);
  }

  if (kind === "paragraph") {
    return await renderParagraph(
      element.props as unknown as ParagraphProps,
      context,
      frame,
      id,
      when,
    );
  }

  if (kind === "image") {
    return await renderImage(element.props as unknown as ImageProps, context, frame, id, when);
  }

  if (kind === "graph") {
    return await renderGraph(element.props as unknown as GraphProps, context, frame, id, when);
  }

  if (kind === "repeat") {
    return await renderRepeat(element.props as unknown as RepeatProps, context, frame, id, when);
  }

  const tocProps = element.props as unknown as TableOfContentsProps;

  return prune<TableOfContentsNode>({
    id,
    kind: "tableOfContents",
    title: await text(tocProps.title, context),
    when,
  });
}

/**
 * Records the derivers a node declared, so the engine runs them per document.
 *
 * They have already run if this build had real data, but the node still has to
 * say it wants them: a published document is read by an engine that was not
 * here for the build. A loop walked into its passes is the exception — it has
 * become several nodes, and the derivers ran once, for all of them.
 */
function withDerivers(
  rendered: DocumentNode | DocumentNode[],
  derivers: CommonElementProps["derivers"],
): DocumentNode | DocumentNode[] {
  if (!derivers || derivers.length === 0 || Array.isArray(rendered)) {
    return rendered;
  }

  return { ...rendered, derivers };
}

async function renderSection(
  props: SectionProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<SectionNode> {
  const section: SectionNode = {
    id,
    kind: "section",
    title: await requiredText(props.title, context),
    when,
    children: await renderYield(props.children, context, {
      path: `${frame.path}/${id}`,
      idSuffix: frame.idSuffix,
    }),
  };

  return prune(section);
}

/**
 * A loop, kept as a loop.
 *
 * With real data the collection is known, so the body is walked now and the
 * entries become ordinary nodes — a preview shows the repetition rather than a
 * description of it. Publishing cannot do that, because the number of entries
 * belongs to a request nobody has made, so the body is published once and the
 * engine walks it.
 */
async function renderRepeat(
  props: RepeatProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<DocumentNode | DocumentNode[]> {
  const as = props.as ?? "item";
  const indexAs = props.indexAs ?? "index";

  if (context.branchMode === "publish") {
    const repeat: RepeatNode = {
      id,
      kind: "repeat",
      source: { scope: "data", path: props.over },
      as,
      indexAs,
      when,
      children: await renderYield(props.children, context, {
        path: `${frame.path}/${id}`,
        idSuffix: frame.idSuffix,
      }),
    };

    return prune(repeat);
  }

  const entries = await context.state.dataProvider.get(props.over);

  if (!Array.isArray(entries)) {
    throw new Error(
      `<Repeat over="${props.over}"> found ${
        entries === undefined ? "nothing" : typeof entries
      } instead of a collection.`,
    );
  }

  const nodes: DocumentNode[] = [];
  const previousItem = context.state.ctx[as];
  const previousIndex = context.state.ctx[indexAs];

  for (const [index, entry] of entries.entries()) {
    context.state.ctx[as] = entry;
    context.state.ctx[indexAs] = index;
    nodes.push(
      ...await renderYield(props.children, context, {
        path: `${frame.path}/${id}/${index}`,
        when,
        idSuffix: frame.idSuffix === undefined ? String(index) : `${frame.idSuffix}-${index}`,
      }),
    );
  }

  context.state.ctx[as] = previousItem;
  context.state.ctx[indexAs] = previousIndex;

  return nodes;
}

async function renderParagraph(
  props: ParagraphProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<ParagraphNode> {
  const body = props.text ?? joinText(props.children, frame);
  const prompts = settled(body, props, frame, id);

  if (!isDynamic(prompts)) {
    const node: ParagraphNode = {
      id,
      kind: "paragraph",
      mode: "static",
      when,
      text: await requiredText(body, context),
    };

    return prune(node);
  }

  if (context.dynamicMode === "placeholder") {
    return prune({
      id,
      kind: "paragraph",
      mode: "dynamic",
      when,
      text: await placeholderText(prompts, id, context),
    });
  }

  const specs = await promptSpecs(prompts, context);
  const node: ParagraphNode = { id, kind: "paragraph", mode: "dynamic", when, prompts: specs };

  if (!context.aiClient) {
    throw new Error(`Dynamic paragraph "${id}" requires an aiClient.`);
  }

  return prune({
    ...node,
    text: await context.aiClient.generateParagraph({
      node,
      prompt: formatPromptText(specs),
      prompts: specs,
      state: context.state,
    }),
  });
}

async function renderImage(
  props: ImageProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<ImageNode> {
  const prompts = settled(props.src, props, frame, id);

  if (!isDynamic(prompts)) {
    if (props.src === undefined) {
      throw new Error(
        `<Image id="${id}"> has neither a src nor any prompt, so nothing says what it shows.`,
      );
    }

    return prune({
      id,
      kind: "image",
      mode: "static",
      when,
      path: await requiredText(props.src, context),
      alt: await text(props.alt, context),
      width: props.width,
      height: props.height,
    });
  }

  if (context.dynamicMode === "placeholder") {
    return prune({
      id,
      kind: "image",
      mode: "dynamic",
      when,
      placeholder: await placeholderText(prompts, id, context),
    });
  }

  const specs = await promptSpecs(prompts, context);
  const node: ImageNode = { id, kind: "image", mode: "dynamic", when, prompts: specs };

  if (!context.aiClient?.generateImage) {
    throw new Error(`Dynamic image "${id}" requires an aiClient.generateImage method.`);
  }

  const result = await context.aiClient.generateImage({
    node,
    prompt: formatPromptText(specs),
    prompts: specs,
    state: context.state,
  });

  return prune({ ...node, ...result, path: result.path });
}

async function renderGraph(
  props: GraphProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<GraphNode> {
  const prompts = settled(props.data === undefined ? undefined : "data", props, frame, id);
  const graphType = props.graphType ?? "bar";

  if (!isDynamic(prompts)) {
    if (props.data === undefined) {
      throw new Error(
        `<Graph id="${id}"> has neither data nor any prompt, so nothing says what it plots.`,
      );
    }

    return prune({
      id,
      kind: "graph",
      mode: "static",
      when,
      graphType,
      data: await jsonText(props.data, context) as JsonObject,
      caption: await text(props.caption, context),
    });
  }

  if (context.dynamicMode === "placeholder") {
    return prune({
      id,
      kind: "graph",
      mode: "dynamic",
      when,
      graphType,
      placeholder: await placeholderText(prompts, id, context),
    });
  }

  const specs = await promptSpecs(prompts, context);
  const node: GraphNode = { id, kind: "graph", mode: "dynamic", when, graphType, prompts: specs };

  if (!context.aiClient?.generateGraph) {
    throw new Error(`Dynamic graph "${id}" requires an aiClient.generateGraph method.`);
  }

  const result = await context.aiClient.generateGraph({
    node,
    prompt: formatPromptText(specs),
    prompts: specs,
    state: context.state,
  });

  return prune({
    ...node,
    graphType: result.graphType ?? graphType,
    data: result.data,
    caption: result.caption,
  });
}

function instanceAt(context: RenderContext, path: string): ComponentInstance {
  const existing = context.instances.get(path);

  if (existing) {
    existing.cursor = 0;
    existing.prompts.systemPrompt = undefined;
    existing.prompts.generalPrompt = undefined;
    existing.prompts.infoPrompt = undefined;
    existing.prompts.negativePrompt = undefined;
    existing.prompts.placeholder = undefined;
    return existing;
  }

  const created: ComponentInstance = { path, cells: [], cursor: 0, prompts: {} };
  context.instances.set(path, created);

  return created;
}

/**
 * Settles a node's id.
 *
 * An explicit id is kept, because that is what an engine and a reviewer refer
 * to. Without one the position supplies it, so a branch or a loop does not force
 * anybody to invent names. Two nodes answering to the same id is always a
 * mistake, and it is reported here rather than resolved by whichever came last.
 */
function claimId(
  context: RenderContext,
  explicit: string | undefined,
  frame: Frame,
  kind: string,
): string {
  const base = explicit ?? `${frame.path.replace(/^\//, "").replaceAll("/", "-") || kind}`;
  const id = frame.idSuffix === undefined ? base : `${base}-${frame.idSuffix}`;
  const owner = context.usedIds.get(id);

  if (owner !== undefined && explicit !== undefined) {
    throw new Error(
      `Two nodes claim the id "${id}" — one at ${owner}, one at ${describe(frame)}. ` +
        "Ids name a node for the engine, so they have to be unique.",
    );
  }

  const unique = owner === undefined ? id : uniqueId(context, id);
  context.usedIds.set(unique, describe(frame));

  return unique;
}

function uniqueId(context: RenderContext, id: string): string {
  let suffix = 2;

  while (context.usedIds.has(`${id}-${suffix}`)) {
    suffix += 1;
  }

  return `${id}-${suffix}`;
}

/**
 * Gathers the conditions a node sits under.
 *
 * Branches nest, so a node can be selected by more than one decision, and the
 * engine has to agree with all of them. Flattening as they combine keeps the
 * published condition a list rather than a chain of pairs.
 */
function allOf(...conditions: Array<Condition | undefined>): Condition | undefined {
  const present = conditions.filter((condition): condition is Condition => condition !== undefined);

  if (present.length <= 1) {
    return present[0];
  }

  return {
    type: "and",
    conditions: present.flatMap((condition) =>
      condition.type === "and" ? condition.conditions : [condition]
    ),
  };
}

function childFrame(frame: Frame, index: number): Frame {
  return { ...frame, path: `${frame.path}/${index}` };
}

function describe(frame: Frame): string {
  return frame.path === "" ? "the document root" : frame.path;
}

function mergePrompts(...drafts: Array<PromptDraft | undefined>): PromptDraft {
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
function settled(
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

function isDynamic(prompts: PromptProps): boolean {
  return prompts.systemPrompt !== undefined || prompts.generalPrompt !== undefined ||
    prompts.infoPrompt !== undefined || prompts.negativePrompt !== undefined;
}

async function promptSpecs(
  prompts: PromptDraft,
  context: RenderContext,
): Promise<PromptSpec[]> {
  const order: PromptKind[] = ["system", "general", "info", "negative"];
  const specs: PromptSpec[] = [];

  for (const kind of order) {
    const value = prompts[promptPropByKind[kind]];

    if (value !== undefined && value.trim() !== "") {
      specs.push({ kind, text: await requiredText(value, context) });
    }
  }

  return specs;
}

async function placeholderText(
  prompts: PromptDraft,
  id: string,
  context: RenderContext,
): Promise<string> {
  const placeholder = prompts.placeholder;

  return placeholder && placeholder.trim() !== ""
    ? await requiredText(placeholder, context)
    : `[Dynamic placeholder: ${id}]`;
}

function formatPromptText(prompts: PromptSpec[]): string {
  return prompts.map((entry) => `${entry.kind.toUpperCase()}: ${entry.text}`).join("\n");
}

function joinText(children: Yield, frame: Frame): string {
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

async function runNodeDerivers(
  props: CommonElementProps,
  context: RenderContext,
): Promise<void> {
  if (context.deriverMode === "preserve") {
    return;
  }

  await runDerivers(props.derivers, context.state, context.derivers);
}

async function text(
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

async function requiredText(value: string, context: RenderContext): Promise<string> {
  return await text(value, context) ?? "";
}

async function jsonText(value: unknown, context: RenderContext): Promise<unknown> {
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

/** Keeps undefined fields out of the published JSON. */
function prune<TNode extends object>(node: TNode): TNode {
  return Object.fromEntries(
    Object.entries(node).filter(([, value]) => value !== undefined),
  ) as TNode;
}
