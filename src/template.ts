import {
  doc,
  type DocumentOptions,
  type DocumentTemplate,
  type NodeChildren,
  type NodeComponent,
  section,
  type TemplateNodeComponent,
} from "./components.ts";
import type { SectionNode } from "./domain/types.ts";

export type TemplateElement<TData = unknown> =
  | DocumentTemplate<TData>
  | NodeComponent<TData>
  | TemplateNodeComponent<TData>
  | TemplateElement<TData>[]
  | false
  | null
  | undefined;

export interface DocumentProps<TData = unknown> extends DocumentOptions {
  children?: TemplateElement<TData>;
}

export type TemplateOptions<TProps, TOptions> = TOptions | ((props: TProps) => TOptions);

export type DocumentComponent<TData, TProps extends Record<string, unknown> = Record<never, never>> =
  (props: TProps & { children?: TemplateElement<TData> }) => DocumentTemplate<TData>;

export type SectionComponent<
  TData,
  TProps extends Record<string, unknown> = Record<never, never>,
> = (props: TProps & { children?: TemplateElement<TData> }) => NodeComponent<TData, SectionNode>;

export function Document<TData = unknown>(props: DocumentProps<TData>): DocumentTemplate<TData> {
  return doc(
    {
      id: props.id,
      title: props.title,
      metadata: props.metadata,
    },
    nodeChildrenFromTemplate(props.children),
  );
}

export function defineDocumentComponent<
  TData,
  TProps extends Record<string, unknown> = Record<never, never>,
>(
  options: TemplateOptions<TProps, DocumentOptions>,
): DocumentComponent<TData, TProps> {
  return (props) =>
    Document<TData>({
      ...resolveTemplateOptions(options, props),
      children: props.children,
    });
}

export interface SectionProps<TData = unknown> {
  id: string;
  title: string;
  children?: TemplateElement<TData>;
}

export function Section<TData = unknown>(
  props: SectionProps<TData>,
): NodeComponent<TData, SectionNode> {
  return section(
    {
      id: props.id,
      title: props.title,
    },
    nodeChildrenFromTemplate(props.children),
  );
}

export function defineSectionComponent<
  TData,
  TProps extends Record<string, unknown> = Record<never, never>,
>(
  options: TemplateOptions<TProps, Omit<SectionProps<TData>, "children">>,
): SectionComponent<TData, TProps> {
  return (props) =>
    Section<TData>({
      ...resolveTemplateOptions(options, props),
      children: props.children,
    });
}

export function Node<TData = unknown>(
  props: { component: NodeComponent<TData> },
): TemplateElement {
  return props.component as unknown as NodeComponent<unknown>;
}

export function template<TData>(element: TemplateElement<TData>): DocumentTemplate<TData> {
  const values = flattenTemplate(element).filter((value) => value !== undefined);

  if (values.length === 1 && isDocumentTemplate(values[0])) {
    return values[0];
  }

  throw new Error("Expected a single Document element");
}

function resolveTemplateOptions<TProps, TOptions>(
  options: TemplateOptions<TProps, TOptions>,
  props: TProps,
): TOptions {
  return typeof options === "function" ? (options as (props: TProps) => TOptions)(props) : options;
}

function nodeChildrenFromTemplate<TData>(
  element: TemplateElement<TData>,
): NodeChildren<TData> {
  const children: NodeChildren<TData> = [];

  for (const value of flattenTemplate(element)) {
    if (value === undefined) {
      continue;
    }

    if (isDocumentTemplate(value)) {
      throw new Error("Document elements cannot be nested inside node children");
    }

    children.push(value);
  }

  return children;
}

function flattenTemplate<TData>(
  element: TemplateElement<TData>,
): Array<DocumentTemplate<TData> | NodeComponent<TData> | undefined> {
  if (element === false || element === null || element === undefined) {
    return [undefined];
  }

  if (Array.isArray(element)) {
    return element.flatMap(flattenTemplate);
  }

  return [element];
}

function isDocumentTemplate<TData>(
  value: DocumentTemplate<TData> | NodeComponent<TData> | TemplateNodeComponent<TData> | undefined,
): value is DocumentTemplate<TData> {
  return Boolean(
    value &&
      typeof value === "object" &&
      "id" in value &&
      "title" in value &&
      "nodes" in value,
  );
}
