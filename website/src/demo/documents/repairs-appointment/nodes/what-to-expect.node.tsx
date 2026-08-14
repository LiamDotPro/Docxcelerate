/** @jsxImportSource docxcelerate/template */
import { Paragraph, useSetPlaceholders, useSetPrompts, useState } from "docxcelerate/template";
import type { RepairsData } from "../types.ts";

export const WhatToExpect: Paragraph = () => {
  const [state] = useState((data: RepairsData) => ({ trade: data.trade }));

  useSetPrompts({
    systemPrompt:
      "You write for social housing residents. Use plain English and short sentences.",
    generalPrompt: `In three short sentences, explain what a ${state.trade} will do ` +
      `during a routine repair visit, how long it usually takes, and what the ` +
      `resident should move out of the way beforehand.`,
    negativePrompt: "Do not repeat the appointment date or the job reference.",
  });

  useSetPlaceholders(`What to expect during a ${state.trade}'s visit.`);

  return <Paragraph id="what-to-expect" />;
};
