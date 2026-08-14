/** @jsxImportSource docxcelerate/template */
import { Paragraph, useSetPlaceholders, useSetPrompts, useState } from "docxcelerate/template";
import type { PolicyData } from "../types.ts";

export const Shopping: Paragraph = () => {
  const [state] = useState((data: PolicyData) => ({ name: data.holderName }));

  useSetPrompts({
    systemPrompt:
      "You write regulated insurance correspondence. Be neutral and never discourage switching.",
    generalPrompt: `Write two sentences reminding ${state.name} that they can compare ` +
      `this renewal against other quotes, and that doing so will not affect ` +
      `their existing cover.`,
    negativePrompt: "Do not quote a price, a percentage, or a competitor name.",
  });

  useSetPlaceholders("A short note on comparing this renewal with other quotes.");

  return <Paragraph id="shopping-around" />;
};
