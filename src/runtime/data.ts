import type { DataProvider, JsonObject } from "../domain/types.ts";
import { getPath } from "./object_path.ts";

/**
 * The simplest place a document's data can come from: an object already in
 * hand.
 *
 * @module
 */

/** A {@linkcode DataProvider} reading dotted paths out of a plain object. */
export class InMemoryDataProvider implements DataProvider {
  /**
   * Wraps an object as a provider.
   *
   * @param values The object to read from. Defaults to an empty one.
   */
  constructor(private readonly values: JsonObject = {}) {}

  /**
   * Reads one dotted path.
   *
   * @param path The path to read, such as `tenant.name`.
   * @returns The value, or `undefined` when the path leads nowhere.
   */
  get(path: string): unknown {
    return getPath(this.values, path);
  }
}
