import type { Condition, DeriverInvocation } from "../domain/types.ts";
import type { ComponentInstance, PromptDraft, RenderContext } from "./context.ts";
import {
  type CommonElementProps,
  isStaticChildren,
  isTemplateElement,
} from "./element.ts";

/**
 * Where a node is, and what it is called.
 *
 * A frame is what a render carries down the tree instead of looking back up
 * it: the path a component is identified by, the conditions gathered from the
 * branches above, and the prompts and derivers a component set on whatever it
 * yields. Naming a node is here too, because a name is the one piece of
 * identity that has to survive the node being moved.
 *
 * @module
 */

export interface Frame {
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

export function instanceAt(context: RenderContext, path: string): ComponentInstance {
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
export function claimId(
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
export function derivedId(props: CommonElementProps, frame: Frame, kind: string): string {
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

export function uniqueId(context: RenderContext, id: string): string {
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
export function allOf(...conditions: Array<Condition | undefined>): Condition | undefined {
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

export function childFrame(frame: Frame, index: number): Frame {
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
export function isLoopPasses(children: readonly unknown[]): boolean {
  if (children.length === 0 || isStaticChildren(children)) {
    return false;
  }

  const first = siblingKey(children[0]);

  return first !== undefined && children.every((child) => siblingKey(child) === first);
}

/** What makes two siblings the same element, written once and walked twice. */
export function siblingKey(child: unknown): string | undefined {
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
export function indexedFrame(frame: Frame, index: number): Frame {
  return {
    ...frame,
    idSuffix: frame.idSuffix === undefined ? String(index) : `${frame.idSuffix}-${index}`,
  };
}

export function describe(frame: Frame): string {
  return frame.path === "" ? "the document root" : frame.path;
}

/** Keeps undefined fields out of the published JSON. */
export function prune<TNode extends object>(node: TNode): TNode {
  return Object.fromEntries(
    Object.entries(node).filter(([, value]) => value !== undefined),
  ) as TNode;
}
