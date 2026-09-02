import type {
  GraphNode,
  ImageNode,
  JsonObject,
  DocumentModel,
  DocumentNode,
  ParagraphNode,
  PromptSpec,
  RepeatNode,
  RuntimeState,
} from "../domain/types.ts";
import { generateGraph, generateImage, generateParagraph } from "./ai_request.ts";
import { evaluateCondition } from "./conditions.ts";
import { createDefaultDeriverRegistry, type DeriverRegistry, runDerivers } from "./derivers.ts";
import { renderTemplate } from "./templates.ts";

/**
 * Settling a published document against real data.
 *
 * This is the engine half of the split: the build left conditions, loops and
 * tokens in place because it could not know the answers, and this is what
 * answers them.
 *
 * @module
 */

/** What {@linkcode resolveDocument} takes beyond the document and its state. */
export interface ResolveOptions {
  /** The derivers the document may invoke. Defaults to the built-in ones. */
  derivers?: DeriverRegistry;
}

/**
 * Walks a published document and settles it: runs derivers, evaluates
 * conditions, unrolls repeats, and substitutes every `{{...}}` token.
 *
 * @param doc The published document.
 * @param state The data, context and clients to resolve against.
 * @param options The derivers the document may invoke.
 * @returns The document with nothing left to decide.
 * @throws If the document invokes a deriver the registry does not hold.
 */
export async function resolveDocument(
  doc: DocumentModel,
  state: RuntimeState,
  options: ResolveOptions = {},
): Promise<DocumentModel> {
  const derivers = options.derivers ?? createDefaultDeriverRegistry();
  const nodes = await resolveNodes(doc.nodes, state, derivers);
  // The furniture is resolved too: a footer naming the recipient's reference
  // is as much a published token as anything in the body, and it is the part
  // that appears on every page if it is left unresolved.
  const header = doc.header ? await resolveNodes(doc.header, state, derivers) : undefined;
  const footer = doc.footer ? await resolveNodes(doc.footer, state, derivers) : undefined;

  return {
    ...doc,
    nodes,
    header,
    footer,
  };
}

async function resolveNodes(
  nodes: DocumentNode[],
  state: RuntimeState,
  derivers: DeriverRegistry,
): Promise<DocumentNode[]> {
  const resolvedNodes: DocumentNode[] = [];

  for (const node of nodes) {
    const resolved = await resolveNode(node, state, derivers);
    if (resolved) {
      resolvedNodes.push(...[resolved].flat());
    }
  }

  return resolvedNodes;
}

async function resolveNode(
  node: DocumentNode,
  state: RuntimeState,
  derivers: DeriverRegistry,
): Promise<DocumentNode | DocumentNode[] | undefined> {
  await runDerivers(node.derivers, state, derivers);

  const conditionMatches = await evaluateCondition(node.when, state);
  if (!conditionMatches) {
    return undefined;
  }

  if (node.kind === "section") {
    return {
      ...node,
      title: node.title ? await renderTemplate(node.title, state) : node.title,
      children: await resolveNodes(node.children, state, derivers),
    };
  }

  if (node.kind === "repeat") {
    return await resolveRepeat(node, state, derivers);
  }

  if (node.kind === "paragraph") {
    return await resolveParagraph(node, state);
  }

  if (node.kind === "image") {
    return await resolveImage(node, state);
  }

  if (node.kind === "graph") {
    return await resolveGraph(node, state);
  }

  // A table, a row and a cell all resolve the same way: they hold nodes and
  // nothing else, so the work is entirely their children's. That is what lets
  // a row carry a condition and a loop produce rows — neither is a case here.
  if (node.kind === "table" || node.kind === "tableRow" || node.kind === "tableCell") {
    return { ...node, children: await resolveNodes(node.children, state, derivers) };
  }

  return node;
}

/**
 * Walks a published loop, now that the request has said how long it is.
 *
 * Each pass binds the entry and its position into `ctx`, so the body reads them
 * the way it reads anything else, and suffixes ids with the index so the nodes
 * of one pass stay distinguishable from the next. The passes are spliced in
 * where the loop stood rather than wrapped in anything, because a wrapper would
 * be a section, and a section is a heading nobody asked for.
 */
async function resolveRepeat(
  node: RepeatNode,
  state: RuntimeState,
  derivers: DeriverRegistry,
): Promise<DocumentNode[]> {
  const entries = await state.dataProvider.get(node.source.path);
  const previousEntry = state.ctx[node.as];
  const previousIndex = state.ctx[node.indexAs];
  const children: DocumentNode[] = [];

  let pass = 0;

  for (const [index, entry] of (Array.isArray(entries) ? entries : []).entries()) {
    state.ctx[node.as] = entry;
    state.ctx[node.indexAs] = index;

    // The test a `.filter()` left behind. Entries that fail it are not walked at
    // all, so they take no pass number either — the passes a recipient sees are
    // numbered from the entries that survived, not from the ones that did not.
    if (node.where && !await evaluateCondition(node.where, state)) {
      continue;
    }

    for (const child of node.children) {
      const resolved = await resolveNode(suffixIds(child, pass), state, derivers);

      if (resolved) {
        children.push(...[resolved].flat());
      }
    }

    pass += 1;
  }

  state.ctx[node.as] = previousEntry;
  state.ctx[node.indexAs] = previousIndex;

  return children;
}

function suffixIds(node: DocumentNode, index: number): DocumentNode {
  const suffixed = { ...node, id: `${node.id}-${index}` };

  // Every kind that holds nodes has to carry the suffix down, or a pass's
  // children keep the ids of the pass before them. A row produced by a loop is
  // the case that made this a list rather than two ifs.
  if (
    suffixed.kind === "section" || suffixed.kind === "repeat" || suffixed.kind === "table" ||
    suffixed.kind === "tableRow" || suffixed.kind === "tableCell"
  ) {
    return { ...suffixed, children: suffixed.children.map((child) => suffixIds(child, index)) };
  }

  return suffixed;
}

async function resolveParagraph(
  node: ParagraphNode,
  state: RuntimeState,
): Promise<ParagraphNode> {
  if (node.mode === "static") {
    return {
      ...node,
      text: node.text ? await renderTemplate(node.text, state) : "",
    };
  }

  const prompts = await resolvePrompts(node.prompts, state);

  return {
    ...node,
    prompts,
    text: await generateParagraph(state.aiClient, node, prompts, state),
  };
}

async function resolveImage(
  node: ImageNode,
  state: RuntimeState,
): Promise<ImageNode> {
  if (node.mode === "static") {
    return {
      ...node,
      path: node.path ? await renderTemplate(node.path, state) : node.path,
      alt: node.alt ? await renderTemplate(node.alt, state) : node.alt,
    };
  }

  const prompts = await resolvePrompts(node.prompts, state);
  const result = await generateImage(state.aiClient, { ...node, prompts }, prompts, state);

  return {
    ...node,
    prompts,
    path: result.path,
    alt: result.alt,
    width: result.width,
    height: result.height,
  };
}

async function resolveGraph(
  node: GraphNode,
  state: RuntimeState,
): Promise<GraphNode> {
  if (node.mode === "static") {
    return {
      ...node,
      data: node.data ? await renderJsonObjectTemplates(node.data, state) : node.data,
      caption: node.caption ? await renderTemplate(node.caption, state) : node.caption,
    };
  }

  const prompts = await resolvePrompts(node.prompts, state);
  const result = await generateGraph(state.aiClient, { ...node, prompts }, prompts, state);

  return {
    ...node,
    prompts,
    graphType: result.graphType ?? node.graphType,
    data: result.data,
    caption: result.caption,
  };
}

async function resolvePrompts(
  prompts: PromptSpec[] | undefined,
  state: RuntimeState,
): Promise<PromptSpec[]> {
  return await Promise.all(
    (prompts ?? []).map(async (prompt) => ({
      ...prompt,
      text: await renderTemplate(prompt.text, state),
    })),
  );
}

async function renderJsonObjectTemplates(
  value: JsonObject,
  state: RuntimeState,
): Promise<JsonObject> {
  return await renderJsonTemplates(value, state) as JsonObject;
}

async function renderJsonTemplates(value: unknown, state: RuntimeState): Promise<unknown> {
  if (typeof value === "string") {
    return await renderTemplate(value, state);
  }

  if (Array.isArray(value)) {
    return await Promise.all(value.map((item) => renderJsonTemplates(item, state)));
  }

  if (value && typeof value === "object") {
    const entries = await Promise.all(
      Object.entries(value).map(async ([key, nestedValue]) => [
        key,
        await renderJsonTemplates(nestedValue, state),
      ]),
    );

    return Object.fromEntries(entries);
  }

  return value;
}
