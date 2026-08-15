import { Paragraph, useState } from "docxcelerate/template";
import type { OfferData } from "../types.ts";

/**
 * Formatting is ordinary TypeScript — no template language, so an Oxford comma
 * and a singular/plural switch cost one line each. It happens in the state
 * initializer, which is where a component does its thinking.
 *
 * Reading the list here settles it during the build. That is right for a
 * document produced from data you hold; a document published to an engine would
 * reach for a deriver instead, so the wording follows a list nobody has yet.
 */
export const Conditions: Paragraph = () => {
  const [state] = useState((data: OfferData) => ({
    list: formatList(data.conditions),
    verb: data.conditions.length === 1 ? "condition is" : "conditions are",
  }));

  return (
    <Paragraph id="conditions-summary">
      This offer is conditional. The {state.verb} {state.list}. We will confirm your place
      automatically once your results reach us.
    </Paragraph>
  );
};

function formatList(items: string[]): string {
  if (items.length <= 1) return items[0] ?? "";
  return `${items.slice(0, -1).join(", ")}, and ${items.at(-1)}`;
}
