import { isTemplateNodeComponent, type TemplateNodeComponent } from "../components.ts";
import type { TemplateElement } from "../template.ts";

export type TemplateComponent<P = Record<string, unknown>> =
  | ((props: P) => TemplateElement)
  | TemplateNodeComponent<unknown>;

export function jsx<P extends Record<string, unknown>>(
  type: TemplateComponent<P>,
  props: P,
): TemplateElement {
  if (typeof type !== "function") {
    throw new Error("Docxcelerate template JSX only supports component functions");
  }

  if (isTemplateNodeComponent(type)) {
    return type;
  }

  return (type as (props: P) => TemplateElement)((props ?? {}) as P);
}

export const jsxs = jsx;

export function Fragment(props: { children?: TemplateElement }): TemplateElement {
  return props.children;
}

// deno-lint-ignore no-namespace
export namespace JSX {
  // deno-lint-ignore no-explicit-any
  export type Element = TemplateElement<any>;
  export interface ElementChildrenAttribute {
    children: Record<string, unknown>;
  }
  // deno-lint-ignore no-empty-interface
  export interface IntrinsicElements {}
}
