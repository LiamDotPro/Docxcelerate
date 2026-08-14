/** @jsxImportSource docxcelerate/template */
import { Paragraph, useSetPlaceholders, useSetPrompts, useState } from "docxcelerate/template";
import type { OfferData } from "../types.ts";

/**
 * The one paragraph worth generating. Everything else in this document is
 * deterministic, so this is the only node that needs an engine at all.
 */
export const TutorNote: Paragraph = () => {
  const [state] = useState((data: OfferData) => ({
    applicantName: data.applicantName,
    interviewer: data.interviewer,
    portfolioTheme: data.portfolioTheme,
  }));

  useSetPrompts({
    systemPrompt:
      "You are an admissions tutor. Be warm but never effusive, and never promise outcomes.",
    generalPrompt: `Write two warm, specific sentences from ${state.interviewer} to ` +
      `${state.applicantName}, referring to their interest in ${state.portfolioTheme}. ` +
      `Mention one thing they should read before term starts.`,
    negativePrompt: "Do not restate the offer, the conditions, or the reply deadline.",
  });

  useSetPlaceholders(
    `A short personal note from ${state.interviewer} about ${state.applicantName}'s interview.`,
  );

  return <Paragraph id="tutor-note" />;
};
