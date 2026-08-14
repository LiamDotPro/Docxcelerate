import {
  type Component,
  createElement,
  type HostElementType,
  type HostKind,
  hostKindOf,
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
 */
export function jsx<P, K extends HostKind>(
  type: HostElementType<P, K>,
  props: P,
  key?: string,
): TemplateElement<K>;
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

export const jsxs = jsx;

export function Fragment(props: { children?: Yield }): Yield {
  return props.children;
}

function describe(value: unknown): string {
  if (typeof value === "string") {
    return `the tag <${value}>`;
  }

  return value === null ? "null" : typeof value;
}

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

  // deno-lint-ignore no-explicit-any
  export type Element = TemplateElement<any>;
  export interface ElementChildrenAttribute {
    children: Record<string, unknown>;
  }
  // deno-lint-ignore no-empty-interface
  export interface IntrinsicElements {}
}
