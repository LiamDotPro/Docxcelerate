import type { DataReference, RuntimeState, ValueExpression } from "../domain/types.ts";
import { getPath } from "./object_path.ts";

const TEMPLATE_REF_PATTERN = /\{\{\s*(data|ctx|derived)\.([A-Za-z0-9_.-]+)\s*\}\}/g;

export async function renderTemplate(source: string, state: RuntimeState): Promise<string> {
  const replacements = await Promise.all(
    Array.from(source.matchAll(TEMPLATE_REF_PATTERN)).map(async (match) => {
      const value = await resolveReference({
        scope: match[1] as DataReference["scope"],
        path: match[2],
      }, state);
      return { token: match[0], value: valueToString(value) };
    }),
  );

  return replacements.reduce(
    (rendered, replacement) => rendered.replaceAll(replacement.token, replacement.value),
    source,
  );
}

export async function resolveValueExpression(
  expression: ValueExpression,
  state: RuntimeState,
): Promise<unknown> {
  if (expression.type === "literal") {
    return expression.value;
  }

  return await resolveReference(expression.ref, state);
}

export async function resolveReference(
  ref: DataReference,
  state: RuntimeState,
): Promise<unknown> {
  if (ref.scope === "data") {
    return await state.dataProvider.get(ref.path);
  }

  if (ref.scope === "ctx") {
    return getPath(state.ctx, ref.path);
  }

  return getPath(state.derived, ref.path);
}

function valueToString(value: unknown): string {
  if (value === undefined || value === null) {
    return "";
  }

  if (typeof value === "string") {
    return value;
  }

  return JSON.stringify(value);
}
