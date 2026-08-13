import { paragraph } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/**
 * The minimum a dynamic paragraph needs: one prompt, and a placeholder so the
 * preview still reads as a letter.
 */
export const NextSteps = paragraph<SampleData>({
  id: "next-steps",
  placeholder: (data) =>
    `Your membership renews automatically on ${data.renewsOn}. ` +
    `Nothing is needed from you unless you want to change plan.`,
  generalPrompt: (data) =>
    `In two sentences, tell ${data.memberName} that their membership renews ` +
    `automatically on ${data.renewsOn} and how to change plan before then.`,
});
