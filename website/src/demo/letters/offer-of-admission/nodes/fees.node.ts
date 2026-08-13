import { paragraph } from "docxcelerate";
import type { OfferData } from "../types.ts";

/**
 * An optional field changes the paragraph rather than leaving a gap — the kind
 * of branch that is awkward in a mail-merge and trivial here.
 */
export const Fees = paragraph<OfferData>({
  id: "fees",
  render: (data) => {
    const base = `You have been assessed as a ${data.feeStatus} fee payer, so ` +
      `tuition for your first year will be ${data.tuitionFee}.`;

    if (!data.scholarship) {
      return `${base} Details of loans, bursaries and hardship funding are in ` +
        `the enclosed funding guide.`;
    }

    return `${base} We are also pleased to award you the ` +
      `${data.scholarship.name}, worth ${data.scholarship.amount} for the ` +
      `duration of your course. No separate application is needed.`;
  },
});
