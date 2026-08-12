import {
  letter,
  type LetterOptions,
  type LetterTemplate,
  type NodeChildren,
  type NodeComponent,
  section,
  type TemplateNodeComponent,
} from "./components.ts";
import type { SectionNode } from "./domain/types.ts";

export type TemplateElement<TData = unknown> =
  | LetterTemplate<TData>
  | NodeComponent<TData>
  | TemplateNodeComponent<TData>
  | TemplateElement<TData>[]
  | false
  | null
  | undefined;

export interface LetterProps<TData = unknown> extends LetterOptions {
  children?: TemplateElement<TData>;
}

export type TemplateOptions<TProps, TOptions> = TOptions | ((props: TProps) => TOptions);

export type LetterComponent<TData, TProps extends Record<string, unknown> = Record<never, never>> =
  (props: TProps & { children?: TemplateElement<TData> }) => LetterTemplate<TData>;

export type SectionComponent<
  TData,
  TProps extends Record<string, unknown> = Record<never, never>,
> = (props: TProps & { children?: TemplateElement<TData> }) => NodeComponent<TData, SectionNode>;

export function Letter<TData = unknown>(props: LetterProps<TData>): LetterTemplate<TData> {
  return letter(
    {
      id: props.id,
      title: props.title,
      metadata: props.metadata,
    },
    nodeChildrenFromTemplate(props.children),
  );
}

export function defineLetterComponent<
  TData,
  TProps extends Record<string, unknown> = Record<never, never>,
>(
  options: TemplateOptions<TProps, LetterOptions>,
): LetterComponent<TData, TProps> {
  return (props) =>
    Letter<TData>({
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

export function template<TData>(element: TemplateElement<TData>): LetterTemplate<TData> {
  const values = flattenTemplate(element).filter((value) => value !== undefined);

  if (values.length === 1 && isLetterTemplate(values[0])) {
    return values[0];
  }

  throw new Error("Expected a single Letter element");
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

    if (isLetterTemplate(value)) {
      throw new Error("Letter elements cannot be nested inside node children");
    }

    children.push(value);
  }

  return children;
}

function flattenTemplate<TData>(
  element: TemplateElement<TData>,
): Array<LetterTemplate<TData> | NodeComponent<TData> | undefined> {
  if (element === false || element === null || element === undefined) {
    return [undefined];
  }

  if (Array.isArray(element)) {
    return element.flatMap(flattenTemplate);
  }

  return [element];
}

function isLetterTemplate<TData>(
  value: LetterTemplate<TData> | NodeComponent<TData> | TemplateNodeComponent<TData> | undefined,
): value is LetterTemplate<TData> {
  return Boolean(
    value &&
      typeof value === "object" &&
      "id" in value &&
      "title" in value &&
      "nodes" in value,
  );
}
