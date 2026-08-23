/**
 * The JSX runtime the compiler emits calls to. You import this by configuring
 * it, not by hand.
 *
 * ```json
 * {
 *   "compilerOptions": {
 *     "jsx": "react-jsx",
 *     "jsxImportSource": "@docxcelerate/docxcelerate/template"
 *   }
 * }
 * ```
 *
 * @module
 */

import {
  type Component,
  createElement,
  type HostElementType,
  type HostKind,
  hostKindOf,
  staticChildrenMarker,
  type TemplateElement,
  type Yield,
} from "./element.ts";

/**
 * Builds an element. It does not run anything.
 *
 * This is the difference that makes components possible. Evaluating JSX used to
 * call the component immediately — at module scope, where no data exists and no
 * decision can be made. Now it records the type and its props, and the renderer
 * calls it later, once, with data in hand and a hook context around it. Every
 * other feature here follows from that.
 *
 * @typeParam P The props the type takes.
 * @typeParam K What the element turns into.
 * @param type A document element such as `Paragraph`.
 * @param props The props written on the tag, children included.
 * @param key Identity across renders, for elements built in a list.
 * @returns The recorded element.
 */
export function jsx<P, K extends HostKind>(
  type: HostElementType<P, K>,
  props: P,
  key?: string,
): TemplateElement<K>;
/**
 * Builds an element from a component. It does not call the component.
 *
 * @typeParam P The props the component takes.
 * @param type A component function.
 * @param props The props written on the tag, children included.
 * @param key Identity across renders, for elements built in a list.
 * @returns The recorded element.
 */
export function jsx<P>(
  type: Component<P>,
  props: P,
  key?: string,
): TemplateElement<"component">;
export function jsx(
  type: unknown,
  props: Record<string, unknown> | null,
  key?: string,
): TemplateElement {
  if (typeof type !== "function") {
    throw new Error(
      "Docxcelerate JSX accepts document elements and component functions, " +
        `not ${describe(type)}. Intrinsic tags like <div> have no meaning in a document.`,
    );
  }

  return createElement(hostKindOf(type) ?? "component", type, props ?? {}, key);
}

/**
 * What the compiler calls for a tag whose children were written out.
 *
 * It marks that array, which is the one thing {@linkcode jsx} cannot know. A
 * `.map()` and a pair of hand-written siblings both arrive as an array; only
 * the second was chosen element by element, so only the second has ids somebody
 * picked and would want a collision reported for.
 */
export const jsxs: typeof jsx = ((type: unknown, props: Record<string, unknown> | null, key?: string) => {
  if (props && Array.isArray(props.children)) {
    Object.defineProperty(props.children, staticChildrenMarker, { value: true });
  }

  return (jsx as (t: unknown, p: Record<string, unknown> | null, k?: string) => TemplateElement)(
    type,
    props,
    key,
  );
}) as typeof jsx;

/**
 * Groups nodes without adding one of its own, written as `<>...</>`.
 *
 * @param props The nodes to group.
 * @returns Those nodes, unchanged.
 */
export function Fragment(props: { children?: Yield }): Yield {
  return props.children;
}

function describe(value: unknown): string {
  if (typeof value === "string") {
    return `the tag <${value}>`;
  }

  return value === null ? "null" : typeof value;
}

/** What TypeScript consults to type-check a tag. */
// deno-lint-ignore no-namespace
export namespace JSX {
  /**
   * What may stand in a tag position.
   *
   * Stated explicitly because a component here does not return an element the
   * way a UI component does: it may return nothing, a list, a string of text,
   * or a promise. Without this, TypeScript would hold every component to
   * returning exactly one element and reject the conditionals that are the
   * point of the model.
   */
  export type ElementType =
    // deno-lint-ignore no-explicit-any
    | HostElementType<any, HostKind>
    // deno-lint-ignore no-explicit-any
    | ((props: any) => Yield | Promise<Yield>);

  /** What a tag evaluates to. */
  // deno-lint-ignore no-explicit-any
  export type Element = TemplateElement<any>;
  /** Names the prop nested tags are collected into. */
  export interface ElementChildrenAttribute {
    /** The prop children arrive on. */
    children: Record<string, unknown>;
  }
  /**
   * Deliberately empty: a document has no intrinsic tags, so `<div>` is a
   * mistake TypeScript can catch rather than something the renderer has to
   * explain later.
   */
  // deno-lint-ignore no-empty-interface
  export interface IntrinsicElements {}
}
