import { paragraph } from "docxcelerate";
import type { RepairsData } from "../types.ts";

export const WhatToExpect = paragraph<RepairsData>({
  id: "what-to-expect",
  placeholder: (data) => `What to expect during a ${data.trade}'s visit.`,
  generalPrompt: (data) =>
    `In three short sentences, explain what a ${data.trade} will do during a ` +
    `routine repair visit, how long it usually takes, and what the resident ` +
    `should move out of the way beforehand.`,
  systemPrompt: () =>
    "You write for social housing residents. Use plain English and short sentences.",
  negativePrompt: () => "Do not repeat the appointment date or the job reference.",
});
