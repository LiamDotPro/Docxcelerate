import type {
  Condition,
  DeriverInvocation,
  DocumentNode,
  GraphNode,
  ImageNode,
  JsonObject,
  PageBreakNode,
  PageNumberNode,
  ParagraphNode,
  RepeatNode,
  SectionNode,
  TableCellNode,
  TableNode,
  TableOfContentsNode,
  TableRowNode,
} from "../domain/types.ts";
import { evaluateCondition, invertCondition } from "../runtime/conditions.ts";
import { runDerivers } from "../runtime/derivers.ts";
import type { BranchProps } from "./branch.ts";
import {
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
  CellProps,
  DocumentProps,
  GraphProps,
  ImageProps,
  PageNumberProps,
  ParagraphProps,
  RowProps,
  SectionProps,
  TableOfContentsProps,
  TableProps,
} from "./elements.ts";
import type { LoopProps } from "./loop.ts";
import {
  formatPromptText,
  isDynamic,
  joinText,
  jsonText,
  mergePrompts,
  placeholderText,
  promptSpecs,
  requiredText,
  settled,
  text,
} from "./content.ts";
import {
  allOf,
  childFrame,
  claimId,
  describe,
  type Frame,
  indexedFrame,
  instanceAt,
  isLoopPasses,
  prune,
} from "./frame.ts";

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
): Promise<{ header?: DocumentNode[]; footer?: DocumentNode[] }> {
  const header = props.header === undefined
    ? undefined
    : await renderYield(props.header, context, { path: "@header" });
  const footer = props.footer === undefined
    ? undefined
    : await renderYield(props.footer, context, { path: "@footer" });

  return {
    header: header && header.length > 0 ? header : undefined,
    footer: footer && footer.length > 0 ? footer : undefined,
  };
}

async function renderYield(
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

