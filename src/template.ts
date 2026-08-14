import type { DocumentTemplate } from "./components.ts";
import { hostKindOf, isTemplateElement, type TemplateElement } from "./template/element.ts";
import type { DocumentProps } from "./template/elements.ts";

export * from "./template/element.ts";
export * from "./template/elements.ts";
export * from "./template/hooks.ts";
export {
  and,
  branch,
  compare,
  ctxPath,
  dataPath,
  derivedPath,
  literal,
  or,
  refValue,
  truthy,
} from "./template/branch.ts";
export type { BranchMode, DeriverMode, DynamicMode } from "./template/context.ts";
export { createPublishData, isPublishValue } from "./template/publish.ts";
export type { ComponentRuntimeOptions, DocumentTemplate } from "./components.ts";
export { buildDocument } from "./components.ts";

/**
 * Names a document tree so a project can point at it.
 *
 * Nothing here is rendered. Evaluating the JSX only builds elements, and the
 * root's id and title are plain props, so a template can be read for what it is
 * called long before anybody has data to build it with.
 */
export function template<TData>(element: unknown): DocumentTemplate<TData> {
  if (!isTemplateElement(element) || hostKindOf(element.type) !== "document") {
    throw new Error(
      "A template is a single <Document> element. " +
        `Received ${describe(element)}.`,
    );
  }

  const props = element.props as unknown as DocumentProps;

  if (!props.id || !props.title) {
    throw new Error("<Document> needs both an id and a title.");
  }

  return {
    schemaVersion: "docxcelerate.template/v0",
    id: props.id,
    title: props.title,
    metadata: props.metadata,
    element: element as TemplateElement<"document">,
  };
}

function describe(value: unknown): string {
  if (isTemplateElement(value)) {
    const kind = hostKindOf(value.type) ?? value.kind;
    return kind === "component"
      ? "a component — call it from inside <Document>, or return <Document> from it"
      : `a <${kind}> element`;
  }

  return value === null ? "null" : typeof value;
}
