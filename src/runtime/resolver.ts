import type {
  GraphNode,
  ImageNode,
  JsonObject,
  LetterDocument,
  LetterNode,
  ParagraphNode,
  PromptSpec,
  RuntimeState,
} from "../domain/types.ts";
import { evaluateCondition } from "./conditions.ts";
import { createDefaultDeriverRegistry, type DeriverRegistry, runDerivers } from "./derivers.ts";
import { renderTemplate } from "./templates.ts";

export interface ResolveOptions {
  derivers?: DeriverRegistry;
}

export async function resolveLetter(
  letter: LetterDocument,
  state: RuntimeState,
  options: ResolveOptions = {},
): Promise<LetterDocument> {
  const derivers = options.derivers ?? createDefaultDeriverRegistry();
  const nodes = await resolveNodes(letter.nodes, state, derivers);

  return {
    ...letter,
    nodes,
  };
}

async function resolveNodes(
  nodes: LetterNode[],
  state: RuntimeState,
  derivers: DeriverRegistry,
): Promise<LetterNode[]> {
  const resolvedNodes: LetterNode[] = [];

  for (const node of nodes) {
    const resolved = await resolveNode(node, state, derivers);
    if (resolved) {
      resolvedNodes.push(resolved);
    }
  }

  return resolvedNodes;
}

async function resolveNode(
  node: LetterNode,
  state: RuntimeState,
  derivers: DeriverRegistry,
): Promise<LetterNode | undefined> {
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

  if (node.kind === "paragraph") {
    return await resolveParagraph(node, state);
  }

  if (node.kind === "image") {
    return await resolveImage(node, state);
  }

  if (node.kind === "graph") {
    return await resolveGraph(node, state);
  }

  return node;
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
  const prompt = formatPromptText(prompts);

  return {
    ...node,
    prompts,
    text: await state.aiClient.generateParagraph({
      node,
      prompt,
      prompts,
      state,
    }),
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
  const prompt = formatPromptText(prompts);

  if (!state.aiClient.generateImage) {
    throw new Error(`Dynamic image "${node.id}" requires an aiClient.generateImage method.`);
  }

  const result = await state.aiClient.generateImage({
    node: {
      ...node,
      prompts,
    },
    prompt,
    prompts,
    state,
  });

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
  const prompt = formatPromptText(prompts);

  if (!state.aiClient.generateGraph) {
    throw new Error(`Dynamic graph "${node.id}" requires an aiClient.generateGraph method.`);
  }

  const result = await state.aiClient.generateGraph({
    node: {
      ...node,
      prompts,
    },
    prompt,
    prompts,
    state,
  });

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

function formatPromptText(prompts: PromptSpec[]): string {
  return prompts.map((entry) => `${entry.kind.toUpperCase()}: ${entry.text}`).join("\n");
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
