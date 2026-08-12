import type { Condition, RuntimeState } from "../domain/types.ts";
import { resolveReference } from "./templates.ts";

export async function evaluateCondition(
  condition: Condition | undefined,
  state: RuntimeState,
): Promise<boolean> {
  if (!condition) {
    return true;
  }

  const value = await resolveReference(condition.ref, state);

  if (condition.type === "not") {
    return !value;
  }

  return Boolean(value);
}
