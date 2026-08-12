import type { DataProvider, JsonObject } from "../domain/types.ts";
import { getPath } from "./object_path.ts";

export class InMemoryDataProvider implements DataProvider {
  constructor(private readonly values: JsonObject = {}) {}

  get(path: string): unknown {
    return getPath(this.values, path);
  }
}
