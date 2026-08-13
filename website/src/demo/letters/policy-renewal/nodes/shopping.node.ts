import { paragraph } from "docxcelerate";
import type { PolicyData } from "../types.ts";

export const Shopping = paragraph<PolicyData>({
  id: "shopping-around",
  placeholder: () => "A short note on comparing this renewal with other quotes.",
  generalPrompt: (data) =>
    `Write two sentences reminding ${data.holderName} that they can compare ` +
    `this renewal against other quotes, and that doing so will not affect ` +
    `their existing cover.`,
  systemPrompt: () =>
    "You write regulated insurance correspondence. Be neutral and never discourage switching.",
  negativePrompt: () => "Do not quote a price, a percentage, or a competitor name.",
});
