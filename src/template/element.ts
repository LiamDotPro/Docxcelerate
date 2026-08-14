import type { Condition, DeriverInvocation } from "../domain/types.ts";

export const elementMarker = Symbol.for("docxcelerate.element");
export const hostMarker = Symbol.for("docxcelerate.host");

/**
 * What a piece of the tree turns into.
 *
 * `component` is the kind of an element whose type is a function you wrote:
 * what it yields is not known until it runs, which is the whole point. `branch`
 * and `fragment` are structural and produce no node of their own.
 */
export type HostKind =
  | "document"
  | "section"
  | "paragraph"
  | "image"
  | "graph"
  | "tableOfContents"
  | "repeat"
  | "branch"
  | "fragment";

export type YieldKind = HostKind | "component";

export interface TemplateElement<K extends YieldKind = YieldKind> {
  readonly [elementMarker]: true;
  readonly kind: K;
  readonly type: unknown;
  readonly props: Record<string, unknown>;
  readonly key?: string;
}

export type Falsy = false | null | undefined;

/**
 * Everything a component may hand back.
 *
 * A component element is accepted wherever any kind is, because a function's
 * yield is unknown at the type level — the renderer checks the real one and
 * says so by name if it is wrong. Naming a kind still rejects the mistake worth
 * catching statically: returning a `<Section>` where a paragraph belongs.
 */
export type Yield<K extends YieldKind = YieldKind> =
  | TemplateElement<K | "component" | "fragment" | "branch">
  | string
  | number
  | Falsy
  | ReadonlyArray<Yield<K>>;

export interface HostElementType<P, K extends HostKind> {
  (props: P): TemplateElement<K>;
  readonly [hostMarker]: K;
}

export type Component<P = Record<never, never>, K extends YieldKind = YieldKind> = (
  props: P,
) => Yield<K> | Promise<Yield<K>>;

/** Props every host element understands, whoever wrote them. */
export interface CommonElementProps {
  id?: string;
  /**
   * A decision left to the engine. Written by the branch compiler rather than
   * by hand — an `if` in a component is the authoring surface for this.
   */
  when?: Condition;
  /**
   * Derivers the engine runs before this node, as opposed to `useDeriver`,
   * which runs one now. The difference is when: these survive publishing and
   * run per document, against data this build never sees.
   */
  derivers?: DeriverInvocation[];
}

export function isTemplateElement(value: unknown): value is TemplateElement {
  return Boolean(
    value && typeof value === "object" && (value as Record<symbol, unknown>)[elementMarker] === true,
  );
}

export function isHostElementType(value: unknown): value is HostElementType<unknown, HostKind> {
  return typeof value === "function" &&
    typeof (value as unknown as Record<symbol, unknown>)[hostMarker] === "string";
}

export function hostKindOf(value: unknown): HostKind | undefined {
  return isHostElementType(value)
    ? (value as unknown as Record<symbol, HostKind>)[hostMarker]
    : undefined;
}

/**
 * Declares a host element.
 *
 * The returned function is never called — `jsx` stores it as the element's type
 * and the renderer switches on its kind. It is a function only so that TypeScript
 * accepts it in JSX position, and it throws if something calls it directly,
 * which only happens when a component is invoked by hand instead of returned.
 */
export function host<P, K extends HostKind>(kind: K, name: string): HostElementType<P, K> {
  const element = ((): never => {
    throw new Error(
      `<${name}> is a document element and cannot be called directly. Return it from a component.`,
    );
  }) as unknown as HostElementType<P, K>;

  Object.defineProperty(element, "name", { value: name });
  Object.defineProperty(element, hostMarker, { value: kind });

  return element;
}

export function createElement(
  kind: YieldKind,
  type: unknown,
  props: Record<string, unknown>,
  key?: string,
): TemplateElement {
  return {
    [elementMarker]: true,
    kind,
    type,
    props,
    key,
  } as TemplateElement;
}
