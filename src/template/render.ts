import type {
  Condition,
  DeriverInvocation,
  DocumentNode,
  GraphNode,
  ImageNode,
  InlineImage,
  JsonObject,
  PageBreakNode,
  PageNumberNode,
  ParagraphNode,
  PromptKind,
  PromptSpec,
  RepeatNode,
  SectionNode,
  TableCellNode,
  TableNode,
  TableOfContentsNode,
  TableRowNode,
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
  isStaticChildren,
  isTemplateElement,
  type TemplateElement,
  type Yield,
} from "./element.ts";
import type {
  CellProps,
  DocumentProps,
  GraphProps,
  ImageProps,
  PageNumberProps,
  ParagraphProps,
  PromptProps,
  RowProps,
  SectionProps,
  TableOfContentsProps,
  TableProps,
} from "./elements.ts";
import type { LoopProps } from "./loop.ts";

interface Frame {
  /** Where this sits in the tree. Identity for hooks, and a fallback id. */
  readonly path: string;
  /** Conditions gathered from branches above, carried onto published nodes. */
  readonly when?: Condition;
  /** Prompts set by the component that yielded this element. */
  readonly prompts?: PromptDraft;
  /** The component this came out of, which is what names the node it yields. */
  readonly componentName?: string;
  /** Derivers a component asked for, to be carried onto the node it yields. */
  readonly derivers?: readonly DeriverInvocation[];
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

/**
 * The running header and footer, rendered apart from the body.
 *
 * Their own paths, so the ids they claim do not collide with the body's, and
 * so a component in the footer gets its own hook state rather than sharing
 * with whatever happens to sit first in the document.
 */
export async function renderDocumentFurniture(
  props: DocumentProps,
  context: RenderContext,
): Promise<{
  header?: DocumentNode[];
  footer?: DocumentNode[];
  firstHeader?: DocumentNode[];
  firstFooter?: DocumentNode[];
}> {
  const header = props.header === undefined
    ? undefined
    : await renderYield(props.header, context, { path: "@header" });
  const footer = props.footer === undefined
    ? undefined
    : await renderYield(props.footer, context, { path: "@footer" });

  // First-page furniture distinguishes `false` from absent: `false` is a
  // statement — the first page shows nothing where the others show the running
  // strip — so it becomes an empty array the packer turns into an empty part.
  // Absent means the first page is like every other, and nothing is carried.
  const firstHeader = props.firstHeader === undefined
    ? undefined
    : props.firstHeader === false
    ? []
    : await renderYield(props.firstHeader, context, { path: "@firstHeader" });
  const firstFooter = props.firstFooter === undefined
    ? undefined
    : props.firstFooter === false
    ? []
    : await renderYield(props.firstFooter, context, { path: "@firstFooter" });

  return {
    header: header && header.length > 0 ? header : undefined,
    footer: footer && footer.length > 0 ? footer : undefined,
    firstHeader,
    firstFooter,
  };
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
    const passes = isLoopPasses(resolved);

    for (const [index, child] of resolved.entries()) {
      const childsFrame = childFrame(frame, index);

      nodes.push(...await renderYield(
        child,
        context,
        passes ? indexedFrame(childsFrame, index) : childsFrame,
      ));
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
    componentName: name,
    derivers: context.deriverMode === "preserve"
      ? [...(frame.derivers ?? []), ...instance.derivers]
      : frame.derivers,
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
  // A condition the build already settled has no request in it, so there is
  // nothing for the engine to decide and nothing to publish. Taking the arm now
  // is what an `if` on a local variable is supposed to mean, in either mode.
  if (context.branchMode === "decide" || typeof props.condition === "boolean") {
    const taken = typeof props.condition === "boolean"
      ? props.condition
      : await evaluateCondition(props.condition, context.state);
    const arm = taken ? props.whenTrue : props.whenFalse;

    return arm ? await renderYield(arm(), context, childFrame(frame, taken ? 0 : 1)) : [];
  }

  const condition = props.condition;

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
      when: allOf(frame.when, condition),
    }));
  }

  if (props.whenFalse) {
    nodes.push(...await renderYield(props.whenFalse(), context, {
      ...childFrame(frame, 1),
      when: allOf(frame.when, invertCondition(condition)),
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
  const id = claimId(context, props, frame, kind ?? "node");
  const when = allOf(frame.when, props.when);

  await runNodeDerivers(props, context);

  return withDerivers(
    withVariant(await renderHostKind(kind, element, context, frame, id, when), props.variant),
    [...(frame.derivers ?? []), ...(props.derivers ?? [])],
  );
}

/**
 * Carries a node's variant onto it, whatever kind it turned out to be.
 *
 * Done once here rather than in each kind's renderer, because the variant says
 * nothing about the kind — it is a name the style looks up, and every node can
 * carry one.
 */
function withVariant(
  rendered: DocumentNode | DocumentNode[],
  variant: string | undefined,
): DocumentNode | DocumentNode[] {
  if (variant === undefined || Array.isArray(rendered)) {
    return rendered;
  }

  return { ...rendered, variant };
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
    return await renderLoop(element.props as unknown as LoopProps, context, frame, id, when);
  }

  if (kind === "table") {
    return await renderTable(element.props as unknown as TableProps, context, frame, id, when);
  }

  if (kind === "tableRow") {
    return await renderRow(element.props as unknown as RowProps, context, frame, id, when);
  }

  if (kind === "tableCell") {
    return await renderCell(element.props as unknown as CellProps, context, frame, id, when);
  }

  if (kind === "pageBreak") {
    return prune<PageBreakNode>({ id, kind: "pageBreak", when });
  }

  if (kind === "pageNumber") {
    const pageProps = element.props as unknown as PageNumberProps;

    return prune<PageNumberNode>({
      id,
      kind: "pageNumber",
      format: pageProps.format,
      separator: pageProps.separator,
      when,
    });
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
  derivers: readonly DeriverInvocation[],
): DocumentNode | DocumentNode[] {
  if (!derivers || derivers.length === 0 || Array.isArray(rendered)) {
    return rendered;
  }

  return { ...rendered, derivers: [...derivers] };
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
    // Only `false` is worth carrying: absent means printed, and `prune` drops
    // `undefined` while keeping `false` — so the model says only what departs
    // from the default.
    showTitle: props.showTitle === false ? false : undefined,
    when,
    children: await renderYield(props.children, context, {
      path: `${frame.path}/${id}`,
      idSuffix: frame.idSuffix,
    }),
  };

  return prune(section);
}

/**
 * A table, whose rows are ordinary nodes.
 *
 * The columns are the only thing the table itself carries. Everything else is
 * a child, which is what makes a loop of rows work without the table knowing
 * anything about loops: `.map()` over the lines yields rows, and whether those
 * became nodes here or a `repeat` for the engine to walk is settled by the
 * same code that settles it anywhere else.
 */
async function renderTable(
  props: TableProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<TableNode> {
  const table: TableNode = {
    id,
    kind: "table",
    columns: props.columns.map((column) => prune({ ...column })),
    when,
    children: await renderYield(props.children, context, {
      path: `${frame.path}/${id}`,
      idSuffix: frame.idSuffix,
    }),
  };

  return prune(table);
}

async function renderRow(
  props: RowProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<TableRowNode> {
  const row: TableRowNode = {
    id,
    kind: "tableRow",
    header: props.header ? true : undefined,
    when,
    children: await renderYield(props.children, context, {
      path: `${frame.path}/${id}`,
      idSuffix: frame.idSuffix,
    }),
  };

  return prune(row);
}

/**
 * A cell, which holds text directly or nodes when a line is not enough.
 *
 * `<Cell>{line.qty}</Cell>` is what most cells are, so text goes straight in
 * and becomes the paragraph it would have had to be written as. A cell holding
 * elements is taken at its word instead — that is the description above a
 * muted note, and wrapping it would flatten the two into one line.
 */
async function renderCell(
  props: CellProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<TableCellNode> {
  const childFrame = { path: `${frame.path}/${id}`, idSuffix: frame.idSuffix };
  const cell: TableCellNode = {
    id,
    kind: "tableCell",
    span: props.span,
    align: props.align,
    when,
    children: holdsElements(props.children)
      ? await renderYield(props.children, context, childFrame)
      : await cellText(props.children, context, childFrame, id),
  };

  return prune(cell);
}

/** The single paragraph a cell's text becomes. */
async function cellText(
  children: Yield,
  context: RenderContext,
  frame: Frame,
  id: string,
): Promise<DocumentNode[]> {
  const body = joinText(children, frame);

  if (body === "") {
    return [];
  }

  return [prune<ParagraphNode>({
    id: `${id}-text`,
    kind: "paragraph",
    mode: "static",
    text: await requiredText(body, context),
  })];
}

/** Whether a cell was given nodes rather than something to print. */
function holdsElements(children: Yield): boolean {
  let found = false;

  const walk = (value: Yield): void => {
    if (found || value === false || value === null || value === undefined) {
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(walk);
      return;
    }

    if (isTemplateElement(value)) {
      found = true;
    }
  };

  walk(children);

  return found;
}

/**
 * A loop, kept as a loop.
 *
 * Only a publish build ever reaches this. Building against real data walks the
 * collection with the `.map()` in the standard library, and the entries become
 * ordinary nodes — a preview shows the repetition rather than a description of
 * it. Publishing cannot, because the number of entries belongs to a request
 * nobody has made, so the stand-in intercepts `.map()` and the body it walked
 * once arrives here to be published as the loop the engine walks.
 */
async function renderLoop(
  props: LoopProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<DocumentNode> {
  const loop: RepeatNode = {
    id,
    kind: "repeat",
    source: { scope: props.overScope, path: props.over },
    as: props.as,
    indexAs: props.indexAs,
    where: props.where,
    when,
    children: await renderYield(props.children, context, {
      path: `${frame.path}/${id}`,
      idSuffix: frame.idSuffix,
    }),
  };

  return prune(loop);
}

async function renderParagraph(
  props: ParagraphProps,
  context: RenderContext,
  frame: Frame,
  id: string,
  when: Condition | undefined,
): Promise<ParagraphNode> {
  const inlineElements: InlineImageElement[] = [];
  const body = props.text ?? joinText(props.children, frame, inlineElements);

  // Pictures are rendered as the nodes they are, then carried on the
  // paragraph rather than emitted beside it.
  const inlineImages: InlineImage[] = [];
  for (const found of inlineElements) {
    const image = await renderImage(
      found.element.props as unknown as ImageProps,
      context,
      frame,
      claimId(context, found.element.props as CommonElementProps, frame, "image"),
      undefined,
    );
    inlineImages.push({ at: found.at, image });
  }
  const prompts = settled(body, props, frame, id);

  if (!isDynamic(prompts)) {
    const node: ParagraphNode = {
      id,
      kind: "paragraph",
      mode: "static",
      when,
      align: props.align,
      text: await requiredText(body, context),
      inlineImages: inlineImages.length === 0 ? undefined : inlineImages,
    };

    return prune(node);
  }

  if (context.dynamicMode === "placeholder") {
    return prune({
      id,
      kind: "paragraph",
      mode: "dynamic",
      when,
      align: props.align,
      text: await placeholderText(prompts, id, context),
    });
  }

  const specs = await promptSpecs(prompts, context);
  const node: ParagraphNode = {
    id,
    kind: "paragraph",
    mode: "dynamic",
    when,
    align: props.align,
    prompts: specs,
  };

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
      fallbackPath: await text(props.fallbackSrc, context),
      alt: await text(props.alt, context),
      width: props.width,
      height: props.height,
    });
  }

  if (context.dynamicMode === "placeholder") {
    // The size and the description belong to the node, not to the picture an
    // engine will draw into it. Dropping them left a preview reserving no
    // room for a picture whose dimensions the template had just stated.
    return prune({
      id,
      kind: "image",
      mode: "dynamic",
      when,
      alt: await text(props.alt, context),
      width: props.width,
      height: props.height,
      placeholder: await placeholderText(prompts, id, context),
    });
  }

  const specs = await promptSpecs(prompts, context);
  const node: ImageNode = {
    id,
    kind: "image",
    mode: "dynamic",
    when,
    alt: await text(props.alt, context),
    width: props.width,
    height: props.height,
    prompts: specs,
  };

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
    existing.prompts.examplePrompt = undefined;
    existing.prompts.placeholder = undefined;
    return existing;
  }

  const created: ComponentInstance = { path, cells: [], cursor: 0, prompts: {},
      derivers: [] };
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
/**
 * Names a node, and makes sure nothing else has that name.
 *
 * An id is an address: an engine targets a node by it, and two builds of the
 * same document line up in a diff by it. So an id nobody wrote still has to be
 * worth having. One taken from where a node sits is not — it changes the moment
 * a paragraph is inserted above it, quietly repointing every address below.
 *
 * The name comes from whatever already says what the node is: the id if one was
 * written, then the heading, then the component that yielded it, and only then
 * the kind. Each of those survives a node being moved, and changes only when
 * somebody deliberately renames something.
 */
function claimId(
  context: RenderContext,
  props: CommonElementProps,
  frame: Frame,
  kind: string,
): string {
  const explicit = props.id;
  const base = explicit ?? derivedId(props, frame, kind);
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

/**
 * The name a node takes when nobody wrote one.
 *
 * A heading is already the human name of the thing it heads, and a component is
 * already named after the node it yields. So a `<Greeting />` becomes `greeting`
 * and a section titled "Fees and funding" becomes `fees-and-funding`, and
 * neither has to be said twice.
 */
function derivedId(props: CommonElementProps, frame: Frame, kind: string): string {
  const title = (props as { title?: unknown }).title;
  const fromTitle = typeof title === "string" ? slug(title) : "";

  return fromTitle || slug(frame.componentName ?? "") || slug(kind) || "node";
}

/**
 * Turns a name people read into one an engine can address.
 *
 * The word boundary in `SignOff` and the one in `Sign off` are the same
 * boundary, so both arrive as `sign-off` and renaming between the two styles
 * leaves the address alone.
 */
function slug(value: string): string {
  return value
    .replace(/([a-z0-9])([A-Z])/g, "$1-$2")
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1-$2")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
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

/**
 * Whether a list of siblings is the passes of a loop.
 *
 * Two things arrive here as the same array: children written out in the source,
 * and a list an expression produced. A repeated id means something different in
 * each. Children were chosen one at a time, so a repeat there is a typo and is
 * reported. A `.map()` is one body written once, so every pass is the same
 * element with the same id, and naming them by position is the only thing that
 * could be meant.
 *
 * Every entry has to match, not merely two of them. A loop over one entry is
 * still a loop, and the engine names its single pass by position — so a build
 * that waited to see a repetition would name that node one thing in a preview
 * and another in the document a recipient gets.
 */
function isLoopPasses(children: readonly unknown[]): boolean {
  if (children.length === 0 || isStaticChildren(children)) {
    return false;
  }

  const first = siblingKey(children[0]);

  return first !== undefined && children.every((child) => siblingKey(child) === first);
}

/** What makes two siblings the same element, written once and walked twice. */
function siblingKey(child: unknown): string | undefined {
  if (!isTemplateElement(child)) {
    return undefined;
  }

  const typeName = typeof child.type === "function" ? child.type.name : String(child.type);

  return `${child.kind}:${typeName}:${(child.props as CommonElementProps).id ?? ""}`;
}

/**
 * Names one pass of a loop, the same way a published loop is walked — so a
 * previewed id and a written one stay the same string.
 */
function indexedFrame(frame: Frame, index: number): Frame {
  return {
    ...frame,
    idSuffix: frame.idSuffix === undefined ? String(index) : `${frame.idSuffix}-${index}`,
  };
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
    prompts.infoPrompt !== undefined || prompts.negativePrompt !== undefined ||
    prompts.examplePrompt !== undefined;
}

async function promptSpecs(
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

/** A picture found among a paragraph's children, and where it was found. */
type InlineImageElement = { at: number; element: TemplateElement };

/**
 * The words a paragraph's children spell, and any pictures set among them.
 *
 * Passing `inlineAt` opts into collecting pictures; without it an element
 * child is still the error it always was.
 */
function joinText(children: Yield, frame: Frame, inlineAt?: InlineImageElement[]): string {
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
      // A picture is the one element that belongs inside a line rather than
      // beside it. Where it sits is remembered as an offset into the text
      // built so far, so the words stay one string and the order survives.
      if (hostKindOf(value.type) === "image") {
        inlineAt?.push({ at: parts.join("").length, element: value });
        return;
      }

      throw new Error(
        `A <Paragraph> at ${describe(frame)} was given an element as a child. ` +
          "A paragraph holds text; put a picture inside it, and anything else beside it.",
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

  await runDerivers(
    props.derivers,
    context.state,
    context.derivers,
    context.deriverMode === "placeholder",
  );
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
