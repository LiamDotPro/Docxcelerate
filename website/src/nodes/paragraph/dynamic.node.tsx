import { Paragraph, useSetPlaceholders, useSetPrompts, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * The minimum a dynamic paragraph needs: one prompt, and a placeholder so the
 * preview still reads as a letter. Setting prompts is what makes the node
 * dynamic — nothing declares a mode.
 */
export const NextSteps: Paragraph = () => {
  const [state] = useState((data: SampleData) => ({
    name: data.memberName,
    renewsOn: data.renewsOn,
  }));

  useSetPrompts({
    generalPrompt:
      `In two sentences, tell ${state.name} that their membership renews ` +
      `automatically on ${state.renewsOn} and how to change plan before then.`,
  });

  useSetPlaceholders(
    `Your membership renews automatically on ${state.renewsOn}. ` +
      `Nothing is needed from you unless you want to change plan.`,
  );

  return <Paragraph id="next-steps" />;
};
