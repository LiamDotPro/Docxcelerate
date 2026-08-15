import {
  Paragraph,
  useAvailableTokens,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";
import { money, type SampleData } from "../sample-data.ts";

/**
 * All four slots: general says what to write, info supplies facts without
 * asking for them back, negative fences off the failure modes, system fixes
 * the voice. `useAvailableTokens` is the budget the build allotted this node.
 */
export const Apology: Paragraph = () => {
  const availableTokens = useAvailableTokens();
  const [state] = useState((data: SampleData) => ({
    name: data.memberName,
    centreName: data.centreName,
    plan: data.plan,
    newPrice: data.newPrice,
  }));

  useSetPrompts({
    systemPrompt:
      `You write for a public leisure centre. Plain British English, second ` +
      `person, no marketing language.`,
    generalPrompt: `Apologise to ${state.name} for the main pool closure and say what ` +
      `is still open. At most ${Math.floor(availableTokens / 4)} words.`,
    infoPrompt: `The main pool at ${state.centreName} is resurfacing until 12 October. ` +
      `The teaching pool, gym and classes are unaffected. Members on ` +
      `${state.plan} paying ${money(state.newPrice)} a year get two guest passes ` +
      `as compensation.`,
    negativePrompt:
      `Do not promise a refund, do not give a reopening date beyond 12 October, ` +
      `and do not restate the price.`,
  });

  useSetPlaceholders(
    `The main pool is closed for resurfacing until 12 October. ` +
      `The teaching pool and all land-based classes are running as normal.`,
  );

  return <Paragraph id="pool-closure" />;
};
