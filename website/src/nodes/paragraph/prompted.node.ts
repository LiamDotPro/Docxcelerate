import { paragraph } from "docxcelerate";
import { money, type SampleData } from "../sample-data.ts";

/**
 * All four slots: general says what to write, info supplies facts without
 * asking for them back, negative fences off the failure modes, system fixes
 * the voice. `availableTokens` is the budget the build allotted this node.
 */
export const Apology = paragraph<SampleData>({
  id: "pool-closure",
  placeholder: () =>
    `The main pool is closed for resurfacing until 12 October. ` +
    `The teaching pool and all land-based classes are running as normal.`,
  systemPrompt: () =>
    `You write for a public leisure centre. Plain British English, second ` +
    `person, no marketing language.`,
  generalPrompt: (data, availableTokens) =>
    `Apologise to ${data.memberName} for the main pool closure and say what ` +
    `is still open. At most ${Math.floor(availableTokens / 4)} words.`,
  infoPrompt: (data) =>
    `The main pool at ${data.centreName} is resurfacing until 12 October. ` +
    `The teaching pool, gym and classes are unaffected. Members on ` +
    `${data.plan} paying ${money(data.newPrice)} a year get two guest passes ` +
    `as compensation.`,
  negativePrompt: () =>
    `Do not promise a refund, do not give a reopening date beyond 12 October, ` +
    `and do not restate the price.`,
});
