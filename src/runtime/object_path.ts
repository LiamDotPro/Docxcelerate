import type { JsonObject } from "../domain/types.ts";

export function getPath(source: JsonObject, path: string): unknown {
  return path.split(".").reduce<unknown>((value, segment) => {
    if (value === null || typeof value !== "object") {
      return undefined;
    }
    return (value as JsonObject)[segment];
  }, source);
}

export function setPath(target: JsonObject, path: string, value: unknown): void {
  const segments = path.split(".");
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
