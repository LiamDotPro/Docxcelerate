/**
 * The stand-in a publish build reads its data through.
 *
 * Every path answers with the token an engine substitutes, so a template can be
 * walked in full without any real data — including the branch arms this build
 * would not have taken.
 *
 * @module
 */

const publishTargetMarker = Symbol.for("docxcelerate.publishTarget");

/**
 * Stand-in data for a build whose values belong to a request that has not
 * happened yet.
 *
 * Publishing renders every arm of every branch, including the ones this build
 * would not have taken, so reading a field must never be what stops it. Any
 * path answers, however deep, and interpolating one yields the token the engine
 * substitutes later.
 *
 * What it refuses is the operation whose answer it cannot honestly give.
 * Comparing or counting a value nobody has yet is a decision being made at the
 * wrong time, and quietly returning `false` or `NaN` would publish a document
 * that is wrong for every recipient. So those throw, and say what to do instead.
 */
export function createPublishData(path: string[] = []): unknown {
  const reference = `{{data.${path.join(".")}}}`;
  const target = Object.assign(() => undefined, { [publishTargetMarker]: true });

  return new Proxy(target, {
    get(_target, property) {
      if (property === publishTargetMarker) {
        return true;
      }

      if (property === Symbol.toPrimitive) {
        return (hint: string) => {
          if (hint === "string" || path.length === 0) {
            return reference;
          }

          throw new Error(publishComparisonMessage(path));
        };
      }

      if (property === "toString" || property === Symbol.toStringTag) {
        return () => reference;
      }

      if (property === "valueOf") {
        return () => reference;
      }

      // Awaiting a value walks `then`. Answering with a proxy would make the
      // await hang forever, so this one field is honestly absent.
      if (property === "then" || property === "catch" || property === "finally") {
        return undefined;
      }

      if (typeof property === "symbol") {
        if (property === Symbol.iterator || property === Symbol.asyncIterator) {
          throw new Error(publishLoopMessage(path));
        }

        return undefined;
      }

      if (collectionMembers.has(property)) {
        throw new Error(publishLoopMessage(path));
      }

      return createPublishData([...path, property]);
    },

    apply() {
      throw new Error(
        `${describePath(path)} was called as a function while publishing. ` +
          "Request data holds values, not behaviour.",
      );
    },

    has() {
      return true;
    },
  });
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

function publishComparisonMessage(path: string[]): string {
  return [
    `${describePath(path)} was read as a number while publishing, where its value is not known yet.`,
    "Work that depends on request data has to reach the engine rather than be settled here.",
    "To decide something, branch on it with an `if` in a component, which is compiled into a",
    "condition the engine evaluates per document. To compute or format something — a total, a",
    `currency, a date — use a deriver, and refer to what it produced: derivers={[derive("name",`,
    `{ output: "label", inputs: [dataRef("${path.join(".")}")] })]} then {{derived.label}} in the text.`,
    "Interpolating the value as text needs neither, and already works.",
  ].join(" ");
}

function publishLoopMessage(path: string[]): string {
  return [
    `${describePath(path)} was iterated while publishing, where its length is not known yet.`,
    "A branch has two arms and both can be published; a loop has as many as the request has entries.",
    `Publish the loop itself instead: <Repeat over="${path.join(".")}">…</Repeat>,`,
    "whose body is written once and walked per entry when a document is written.",
  ].join(" ");
}

function describePath(path: string[]): string {
  return path.length === 0 ? "The request data" : `\`data.${path.join(".")}\``;
}

const collectionMembers = new Set([
  "length",
  "map",
  "forEach",
  "filter",
  "reduce",
  "slice",
  "find",
  "some",
  "every",
  "flatMap",
  "join",
  "entries",
  "keys",
  "values",
  "at",
  "includes",
  "indexOf",
  "sort",
  "concat",
]);
