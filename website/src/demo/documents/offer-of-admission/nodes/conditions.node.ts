import { paragraph } from "docxcelerate";
import type { OfferData } from "../types.ts";

/**
 * Formatting is ordinary TypeScript — no template language, so an Oxford comma
 * and a singular/plural switch cost one line each.
 */
export const Conditions = paragraph<OfferData>({
  id: "conditions",
  render: (data) => {
    const list = formatList(data.conditions);
    const verb = data.conditions.length === 1 ? "condition is" : "conditions are";

    return `This offer is conditional. The ${verb} ${list}. We will confirm ` +
      `your place automatically once your results reach us.`;
  },
});

function formatList(items: string[]): string {
  if (items.length <= 1) return items[0] ?? "";
  return `${items.slice(0, -1).join(", ")}, and ${items.at(-1)}`;
}
