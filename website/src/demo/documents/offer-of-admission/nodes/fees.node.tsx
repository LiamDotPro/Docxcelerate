/** @jsxImportSource docxcelerate/template */
import { Paragraph, useState } from "docxcelerate/template";
import type { OfferData } from "../types.ts";

/**
 * An optional field changes the paragraph rather than leaving a gap — the kind
 * of branch that is awkward in a mail-merge and is an `if` here. Each arm has
 * its own id, so the resolved document says which one this applicant got.
 */
export const Fees: Paragraph = () => {
  const [state] = useState((data: OfferData) => ({
    feeStatus: data.feeStatus,
    tuitionFee: data.tuitionFee,
    scholarship: data.scholarship,
  }));

  const assessment =
    `You have been assessed as a ${state.feeStatus} fee payer, so tuition for ` +
    `your first year will be ${state.tuitionFee}.`;

  if (!state.scholarship) {
    return (
      <Paragraph id="fees">
        {assessment} Details of loans, bursaries and hardship funding are in the enclosed
        funding guide.
      </Paragraph>
    );
  }

  return (
    <Paragraph id="fees-with-scholarship">
      {assessment} We are also pleased to award you the {state.scholarship.name}, worth{" "}
      {state.scholarship.amount} for the duration of your course. No separate application
      is needed.
    </Paragraph>
  );
};
