import type { Condition, DeriverInvocation } from "../domain/types.ts";

/**
 * The element tree JSX builds, and the markers that let anything recognise it.
 *
 * Nothing here renders. An element records a type and its props so the renderer
 * can call components later — once, with data in hand — rather than at module
 * scope where there is nothing to decide with.
 *
 * @module
 */

/**
 * Brands an object as a {@linkcode TemplateElement}.
 *
 * Registered rather than unique, so an element built by one copy of the package
 * is still recognised by another.
 */
export const elementMarker: symbol = Symbol.for("docxcelerate.element");
/** Brands a function as a {@linkcode HostElementType}, and records its kind. */
export const hostMarker: symbol = Symbol.for("docxcelerate.host");
/**
 * Brands the children array of a tag whose children were written out in the
 * source, as opposed to one produced by an expression at runtime.
 *
 * Two sibling paragraphs written by hand and two produced by a `.map()` are
 * the same array by the time the renderer sees them. They are not the same
 * mistake, though: the first pair sharing an id is a typo worth reporting, and
 * the second pair sharing one is simply a loop. This is what tells them apart.
 */
export const staticChildrenMarker: symbol = Symbol.for("docxcelerate.staticChildren");

/**
 * Whether an array is the children of a tag as written, rather than a list an
 * expression produced.
 *
 * @param value The value to test.
 * @returns `true` when the array was written out in the source.
 */
export function isStaticChildren(value: unknown): boolean {
  return Array.isArray(value) &&
    (value as unknown as Record<symbol, unknown>)[staticChildrenMarker] === true;
}

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
  | "table"
  | "tableRow"
  | "tableCell"
  | "tableOfContents"
  | "pageBreak"
  | "pageNumber"
  | "repeat"
  | "branch"
  | "fragment";

/** A {@linkcode HostKind}, or a component whose yield is not known until it runs. */
export type YieldKind = HostKind | "component";

/**
 * One node of the element tree: a type and its props, recorded rather than run.
 *
 * @typeParam K What this element turns into.
 */
export interface TemplateElement<K extends YieldKind = YieldKind> {
  /** Brand marking this as an element. */
  readonly [elementMarker]: true;
  /** What this element turns into. */
  readonly kind: K;
  /** The host element or component this was built from. */
  readonly type: unknown;
  /** The props it was given, children included. */
  readonly props: Record<string, unknown>;
  /** Identity across renders, for elements built in a list. */
  readonly key?: string;
}

/** The values JSX drops rather than renders. */
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

/**
 * A built-in element such as `<Paragraph>`, callable only from JSX.
 *
 * @typeParam P The props it takes.
 * @typeParam K What it turns into.
 */
export interface HostElementType<P, K extends HostKind> {
  /**
   * Never actually called — see {@linkcode host}.
   *
   * @param props The element's props.
   * @returns The element.
   */
  (props: P): TemplateElement<K>;
  /** Brand recording which kind this element is. */
  readonly [hostMarker]: K;
}

/**
 * Something you wrote that yields part of a document.
 *
 * @typeParam P The props it takes.
 * @typeParam K What it is allowed to yield.
 */
export type Component<P = Record<never, never>, K extends YieldKind = YieldKind> = (
  props: P,
) => Yield<K> | Promise<Yield<K>>;

/** Props every host element understands, whoever wrote them. */
export interface CommonElementProps {
  /** Identifier for the node, unique within the document. */
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
  /**
   * Which block style the theme should draw this node in — `"band"`, `"panel"`,
   * `"badge"`. A name for what the node is, never what it looks like: the
   * colours live in the style, so a document restyles without a node changing.
   */
  variant?: string;
}

/**
 * Whether a value is an element of the template tree.
 *
 * @param value The value to test.
 * @returns `true` when it carries the element marker.
 */
export function isTemplateElement(value: unknown): value is TemplateElement {
  return Boolean(
    value && typeof value === "object" && (value as Record<symbol, unknown>)[elementMarker] === true,
  );
}

/**
 * Whether a value is a built-in element rather than a component.
 *
 * @param value The value to test.
 * @returns `true` when it carries the host marker.
 */
export function isHostElementType(value: unknown): value is HostElementType<unknown, HostKind> {
  return typeof value === "function" &&
    typeof (value as unknown as Record<symbol, unknown>)[hostMarker] === "string";
}

/**
 * The kind a built-in element produces.
 *
 * @param value The value to read.
 * @returns Its kind, or `undefined` when it is not a host element.
 */
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
 *
 * @typeParam P The props the element takes.
 * @typeParam K What the element turns into.
 * @param kind What the element turns into.
 * @param name The name it is written under, used in error messages.
 * @returns The element type, ready for JSX position.
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

/**
 * Records an element. This is what the JSX runtime calls; there is rarely a
 * reason to call it by hand.
 *
 * @param kind What the element turns into.
 * @param type The host element or component it was built from.
 * @param props The props it was given.
 * @param key Identity across renders, for elements built in a list.
 * @returns The element.
 */
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
