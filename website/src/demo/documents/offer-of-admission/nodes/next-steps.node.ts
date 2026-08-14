import { paragraph } from "docxcelerate";
import type { OfferData } from "../types.ts";

export const NextSteps = paragraph<OfferData>({
  id: "next-steps",
  render: (data) =>
    `To accept, reply through the applicant portal by ${data.replyBy}. ` +
    `If you would like to visit ${data.college} before deciding, our open ` +
    `afternoons run every Thursday through April and you are very welcome.`,
});
