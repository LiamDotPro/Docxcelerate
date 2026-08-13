import { paragraph } from "docxcelerate";
import type { OfferData } from "../types.ts";

/**
 * The one paragraph worth generating. Everything else in this letter is
 * deterministic, so this is the only node that needs an endpoint at all.
 */
export const TutorNote = paragraph<OfferData>({
  id: "tutor-note",
  placeholder: (data) =>
    `A short personal note from ${data.interviewer} about ${data.applicantName}'s interview.`,
  generalPrompt: (data) =>
    `Write two warm, specific sentences from ${data.interviewer} to ` +
    `${data.applicantName}, referring to their interest in ${data.portfolioTheme}. ` +
    `Mention one thing they should read before term starts.`,
  systemPrompt: () =>
    "You are an admissions tutor. Be warm but never effusive, and never promise outcomes.",
  negativePrompt: () =>
    "Do not restate the offer, the conditions, or the reply deadline.",
});
