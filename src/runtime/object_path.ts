import type { JsonObject } from "../domain/types.ts";

/**
 * Reading and writing the dotted paths a document addresses data by.
 *
 * Both halves of this are reached from a published document rather than from
 * code somebody is looking at: `{{data.a.b}}` in a paragraph becomes a read, and
 * a deriver's `output` becomes a write. A published document is data — it can be
 * stored, copied between systems, and handed to an engine that renders documents
 * for more than one tenant — so neither half may be steered into the parts of an
 * object that were never anybody's data.
 *
 * @module
 */

/**
 * The property names that address the language rather than the data.
 *
 * Writing through `__proto__` reaches `Object.prototype`, where a value lands on
 * every object in the process rather than on the one that was addressed. Reading
 * `constructor` hands back a function nobody stored. A document has no business
 * with either, so both are simply not paths.
 */
const forbiddenSegments = new Set(["__proto__", "constructor", "prototype"]);

/**
 * Reads one dotted path out of a bag of values.
 *
 * Only what is actually there is returned. A path that runs off the end of the
 * data, or that names part of the language rather than part of the data, reads
 * as `undefined` — which is what an absent value already looks like, so nothing
 * downstream needs a second case for it.
 *
 * @param source The bag to read from.
 * @param path A dotted path, such as `tenant.name` or `lines.0.amount`.
 * @returns The value, or `undefined` when there is none.
 */
export function getPath(source: JsonObject, path: string): unknown {
  return path.split(".").reduce<unknown>((value, segment) => {
    if (forbiddenSegments.has(segment)) {
      return undefined;
    }

    // A string knows its own length, and a document counting characters is
    // asking a fair question. Everything else about a primitive is the
    // language's, not the data's.
    if (typeof value === "string") {
      return segment === "length" ? value.length : undefined;
    }

    if (value === null || typeof value !== "object") {
      return undefined;
    }

    return Object.hasOwn(value, segment) ? (value as JsonObject)[segment] : undefined;
  }, source);
}

/**
 * Writes a value at a dotted path, creating the objects along the way.
 *
 * @param target The bag to write into.
 * @param path A dotted path, such as `totals.due`.
 * @param value What to store there.
 * @throws If the path names part of the language rather than part of the data.
 */
export function setPath(target: JsonObject, path: string, value: unknown): void {
  const segments = path.split(".");

  for (const segment of segments) {
    if (forbiddenSegments.has(segment)) {
      throw new Error(
        `"${path}" is not a path a document may write to. ` +
          `"${segment}" addresses the language rather than the data, and writing through it ` +
          "would change every object in the process rather than the one named.",
      );
    }
  }

  let cursor = target;

  for (const segment of segments.slice(0, -1)) {
    const next = cursor[segment];

    if (next === null || typeof next !== "object" || Array.isArray(next)) {
      cursor[segment] = {};
    }

    cursor = cursor[segment] as JsonObject;
  }

  cursor[segments.at(-1) ?? path] = value;
}
