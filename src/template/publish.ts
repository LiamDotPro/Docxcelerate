/**
 * The stand-in a publish build reads its data through.
 *
 * Every path answers with the token an engine substitutes, so a template can be
 * walked in full without any real data — including the branch arms this build
 * would not have taken.
 *
 * The stand-in is also what lets ordinary TypeScript survive publishing. It
 * remembers the path it stands for, so a `.map()` over it can publish a loop, a
 * `.filter()` before that can publish the test the engine applies per entry, and
 * an `if` on it can publish a condition — none of them written by hand.
 *
 * What it still refuses is the operation whose answer it cannot honestly give.
 * Not one of those is a dead end: a deriver runs where the data is, so it can do
 * what this cannot, and its code is testable on its own and shareable between
 * documents in a way a closure inside one component is not.
 *
 * @module
 */

import type { Condition, ReferenceScope, ValueExpression } from "../domain/types.ts";
import { createElement, type TemplateElement, type Yield } from "./element.ts";
import { Loop } from "./loop.ts";

const publishTargetMarker = Symbol.for("docxcelerate.publishTarget");
const publishRefMarker = Symbol.for("docxcelerate.publishRef");

/** Which bag a stand-in stands for, and where within it. */
export interface PublishRef {
  /** The scope an engine reads the value from. */
  scope: ReferenceScope;
  /** The dotted path within that scope. */
  path: string;
}

/** What a stand-in carries beyond the path it stands for. */
interface StandInOptions {
  /**
   * Tests a `.filter()` left behind, for the engine to apply per entry.
   *
   * A filtered collection is no longer addressable as a path — which entries
   * survive is decided per document, so its length and its third entry are not
   * knowable here. Carrying the tests is also what tells the stand-in to stop
   * answering those.
   */
  readonly tests?: readonly Condition[];
  /** Told when a field is read, so a predicate that looked can be told from one that did not. */
  readonly onRead?: () => void;
}

/**
 * Stand-in data for a build whose values belong to a request that has not
 * happened yet.
 *
 * Publishing renders every arm of every branch, including the ones this build
 * would not have taken, so reading a field must never be what stops it. Any
 * path answers, however deep, and interpolating one yields the token the engine
 * substitutes later.
 *
 * @param path Where in the request data this stand-in sits.
 * @returns The stand-in.
 */
export function createPublishData(path: string[] = []): unknown {
  return createStandIn("data", path);
}

/**
 * A stand-in for one entry of a collection walked per document.
 *
 * This is what `.map()` hands its callback while publishing. It reads from
 * `ctx` rather than `data`, because the entry is bound by the loop rather than
 * supplied by the caller.
 *
 * @param path Where in the bound context this stand-in sits.
 * @returns The stand-in.
 */
function createContextStandIn(path: string[]): unknown {
  return createStandIn("ctx", path);
}

/**
 * A stand-in for what a deriver produced, for a build that has not run it.
 *
 * @param path Where under the derived scope this stand-in sits.
 * @returns The stand-in.
 */
export function createDerivedStandIn(path: string[]): unknown {
  return createStandIn("derived", path);
}

function createStandIn(
  scope: ReferenceScope,
  path: string[],
  options: StandInOptions = {},
): unknown {
  const reference = `{{${scope}.${path.join(".")}}}`;
  const target = Object.assign(() => undefined, { [publishTargetMarker]: true });
  const tests = options.tests ?? [];
  const filtered = tests.length > 0;

  return new Proxy(target, {
    get(_target, property) {
      if (property === publishTargetMarker) {
        return true;
      }

      if (property === publishRefMarker) {
        return { scope, path: path.join(".") } satisfies PublishRef;
      }

      if (property === Symbol.toPrimitive) {
        return (hint: string) => {
          options.onRead?.();

          if (filtered) {
            throw new Error(publishFilteredValueMessage(scope, path));
          }

          if (hint === "string" || path.length === 0) {
            return reference;
          }

          throw new Error(publishComparisonMessage(scope, path));
        };
      }

      if (property === "toString" || property === Symbol.toStringTag) {
        options.onRead?.();
        return () => reference;
      }

      if (property === "valueOf") {
        options.onRead?.();
        return () => reference;
      }

      // Awaiting a value walks `then`. Answering with a proxy would make the
      // await hang forever, so this one field is honestly absent.
      if (property === "then" || property === "catch" || property === "finally") {
        return undefined;
      }

      options.onRead?.();

      // The two collection operations whose answer does not depend on the
      // entries. `.map()` writes a body once and the engine walks it per entry;
      // `.filter()` states a test and the engine applies it per entry. Neither
      // needs to know how many entries there are, which is the one thing
      // publishing cannot find out.
      if (property === "map") {
        return (callback: MapCallback) => publishMap(scope, path, tests, callback);
      }

      if (property === "filter") {
        return (predicate: FilterCallback) => publishFilter(scope, path, tests, predicate);
      }

      if (typeof property === "symbol") {
        if (property === Symbol.iterator || property === Symbol.asyncIterator) {
          throw new Error(publishIterationMessage(scope, path));
        }

        return undefined;
      }

      if (deriverMembers.has(property)) {
        throw new Error(publishDeriverMessage(scope, path, property));
      }

      if (filtered) {
        throw new Error(publishFilteredValueMessage(scope, path));
      }

      // Everything else is a path segment, `length` and `0` included. Those two
      // read oddly as paths and resolve exactly right, because the engine reads
      // them off the real array the request brought.
      return createStandIn(scope, [...path, property], { onRead: options.onRead });
    },

    apply() {
      throw new Error(
        `${describePath(scope, path)} was called as a function while publishing. ` +
          "Request data holds values, not behaviour.",
      );
    },

    has() {
      return true;
    },
  });
}

/** What `.map()` is given, matching the callback `Array.prototype.map` takes. */
type MapCallback = (entry: unknown, index: unknown) => Yield;
/** What `.filter()` is given, matching the callback `Array.prototype.filter` takes. */
type FilterCallback = (entry: unknown, index: unknown) => unknown;

/**
 * Publishes a `.map()` as the loop it is.
 *
 * The callback runs once, against a stand-in for one entry, and what it yields
 * becomes the body the engine walks per entry. Nothing here knows how many
 * entries there will be, and nothing needs to — that is the whole reason the
 * loop is published rather than unrolled.
 *
 * The entry is bound under a name taken from the source path, so two loops in
 * one document never collide and the same template always publishes the same
 * tokens. Nobody writes that name: the stand-in produces the tokens itself.
 */
function publishMap(
  scope: ReferenceScope,
  path: string[],
  tests: readonly Condition[],
  callback: MapCallback,
): TemplateElement<"repeat"> {
  const as = path.join("_");
  const indexAs = `${as}_index`;

  return createElement("repeat", Loop, {
    over: path.join("."),
    overScope: scope,
    as,
    indexAs,
    where: everyCondition(tests),
    children: callback(createContextStandIn([as]), createContextStandIn([indexAs])),
  }) as TemplateElement<"repeat">;
}

/**
 * Publishes a `.filter()` as the test the engine applies per entry.
 *
 * The predicate runs once, the same way a mapped body does, and what it hands
 * back says which entries survive. A predicate that reads a field and returns it
 * is a test on that field. One that never reads the entry decides the same way
 * for every entry, so it is settled here and the whole collection stands or
 * falls.
 *
 * A predicate that read the entry and still handed back a plain boolean is the
 * case worth stopping on. JavaScript coerced the value before anything could see
 * which field was meant — `!entry.archived` is already `false` by the time it
 * arrives — so publishing it would silently keep every entry or drop every
 * entry, for every recipient.
 */
function publishFilter(
  scope: ReferenceScope,
  path: string[],
  tests: readonly Condition[],
  predicate: FilterCallback,
): unknown {
  const as = path.join("_");
  let read = false;
  const noticeRead = () => {
    read = true;
  };
  const verdict = predicate(
    createStandIn("ctx", [as], { onRead: noticeRead }),
    createStandIn("ctx", [`${as}_index`], { onRead: noticeRead }),
  );
  const ref = publishRefOf(verdict);

  if (ref) {
    return createStandIn(scope, path, { tests: [...tests, { type: "truthy", ref }] });
  }

  if (typeof verdict === "boolean" && !read) {
    // The same answer for every entry, from a predicate that never looked at
    // one. Keeping all of them is the collection itself; keeping none of them is
    // an empty list, which a later `.map()` walks into nothing, exactly right.
    return verdict ? createStandIn(scope, path, { tests }) : [];
  }

  throw new Error(publishPredicateMessage(scope, path));
}

/**
 * Whether a value is the stand-in a publish build reads its data through.
 *
 * That stand-in answers every path with the token an engine substitutes, so a
 * build can walk a template without any real data. Recognising it is how the
 * rest of the build knows not to treat those answers as values.
 *
 * @param value The value to test.
 * @returns `true` when it is the publish stand-in.
 */
export function isPublishValue(value: unknown): boolean {
  return Boolean(
    value !== null &&
      (typeof value === "object" || typeof value === "function") &&
      (value as Record<symbol, unknown>)[publishTargetMarker] === true,
  );
}

/**
 * The scope and path a stand-in stands for.
 *
 * This is what lets an `if` on request data publish as a condition: the
 * compiled condition hands the value over and the reference comes back, without
 * anything having had to trace where the value came from.
 *
 * @param value The value to read.
 * @returns The reference, or `undefined` when the value is a real one.
 */
export function publishRefOf(value: unknown): PublishRef | undefined {
  if (!isPublishValue(value)) {
    return undefined;
  }

  return (value as Record<symbol, PublishRef>)[publishRefMarker];
}

/**
 * Turns a value into the expression that stands for it.
 *
 * This is the bridge a `.map()` needs. Inside one, `visit.cost` is a real
 * number when the build has data and a stand-in when it is publishing, and the
 * component cannot be asked to know which — that would be the two modes leaking
 * into the document. Handing the value over settles it: a real one travels as
 * a literal, and a stand-in already knows the path it stands for, so it travels
 * as the reference the engine reads per pass.
 *
 * @param value The value to carry, real or standing in for one.
 * @returns The expression to use as a deriver input or a comparison side.
 * @throws If the value is a real one that cannot be written into a document.
 */
export function expr(value: unknown): ValueExpression {
  const ref = publishRefOf(value);

  if (ref) {
    return { type: "ref", ref };
  }

  if (typeof value === "number" && !Number.isFinite(value)) {
    // A published document is JSON, and JSON has no way to write these down:
    // both come back as null, so the value would not survive the trip to the
    // engine. Failing here names the arithmetic that produced it, which is the
    // only place it can still be fixed.
    throw new Error(
      `${Number.isNaN(value) ? "NaN" : String(value)} cannot be carried into a document. ` +
        "A document is stored as JSON, which has no way to write it down — it would arrive " +
        "at the engine as null. Check the arithmetic that produced it, or guard the value " +
        "before it reaches the document.",
    );
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return { type: "literal", value };
  }

  throw new Error(
    `A ${value === null ? "null" : typeof value} cannot be carried into a document as a value. ` +
      "Deriver inputs and comparison sides are single values — a string, a number or a " +
      "boolean. Pass the field you mean rather than the object holding it, or compute it " +
      "in a deriver, which runs where the data is.",
  );
}

/** Joins the tests a chain of `.filter()` calls left behind. */
function everyCondition(conditions: readonly Condition[]): Condition | undefined {
  if (conditions.length === 0) {
    return undefined;
  }

  return conditions.length === 1 ? conditions[0] : { type: "and", conditions: [...conditions] };
}

function publishComparisonMessage(scope: ReferenceScope, path: string[]): string {
  return [
    `${describePath(scope, path)} was read as a number while publishing,`,
    "where its value is not known yet.",
    "Work that depends on request data has to reach the engine rather than be settled here.",
    "To decide something, branch on it with an `if` in a component, which is compiled into a",
    "condition the engine evaluates per document. To compute or format something — a total, a",
    "currency, a date — use a deriver, and refer to what it produced.",
    "Interpolating the value as text needs neither, and already works.",
  ].join(" ");
}

function publishIterationMessage(scope: ReferenceScope, path: string[]): string {
  return [
    `${describePath(scope, path)} was iterated while publishing,`,
    "where its length is not known yet.",
    "A `for` loop has to know how many times to go round before it starts; `.map()` does not,",
    "because its body is written once and walked per entry when a document is written.",
    "Use `.map()` instead.",
  ].join(" ");
}

function publishDeriverMessage(scope: ReferenceScope, path: string[], member: string): string {
  return [
    `\`${member}\` was called on ${describePath(scope, path)} while publishing,`,
    "where its entries are not known yet.",
    "`.map()` and `.filter()` are the two that survive, because neither has to see the entries",
    `— one writes a body per entry, the other states a test. \`${member}\` has to look at them,`,
    "so it belongs in a deriver: the deriver runs where the data is, hands back the value or",
    "the list it worked out, and a `.map()` over what it produced carries on as normal. That",
    "code is also testable on its own and shareable between documents, which a closure inside",
    "one component is not.",
  ].join(" ");
}

function publishFilteredValueMessage(scope: ReferenceScope, path: string[]): string {
  return [
    `A filtered ${describePath(scope, path)} was read as a value while publishing.`,
    "Which entries survive is decided per document, so how many there are and which one is",
    "third are not known here. Map over it to write a body per surviving entry, filter it",
    "further, or use a deriver if you need the entries themselves.",
  ].join(" ");
}

function publishPredicateMessage(scope: ReferenceScope, path: string[]): string {
  return [
    `A \`.filter()\` on ${describePath(scope, path)} read the entry and handed back a plain`,
    "boolean, so what it tested cannot be recovered — JavaScript had already reduced it to",
    "true or false. Publishing that would keep every entry or drop every entry, for every",
    "recipient, without saying so.",
    "Return the field itself to test that it is set: `(line) => line.taxable`. For anything",
    "that has to compare or negate, use a deriver, which runs where the data is and hands back",
    "the list you meant.",
  ].join(" ");
}

function describePath(scope: ReferenceScope, path: string[]): string {
  return path.length === 0 ? "The request data" : `\`${scope}.${path.join(".")}\``;
}

/**
 * The collection operations that have to see the entries before they can answer.
 *
 * Every one of these is a deriver rather than a dead end. `length` and index
 * access are deliberately absent: those are ordinary path segments, and the
 * engine reads them straight off the array the request brought.
 */
const deriverMembers = new Set([
  "forEach",
  "reduce",
  "reduceRight",
  "slice",
  "find",
  "findLast",
  "findIndex",
  "some",
  "every",
  "flat",
  "flatMap",
  "join",
  "entries",
  "keys",
  "values",
  "includes",
  "indexOf",
  "lastIndexOf",
  "sort",
  "toSorted",
  "reverse",
  "toReversed",
  "concat",
  "fill",
  "splice",
  "toSpliced",
]);
