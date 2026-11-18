import { Paragraph, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * Items are paragraphs that say they are items. Written one after another they
 * are one list, and the renderer does the counting — the steps here are
 * numbered without a number being typed, and a bulleted point under a step
 * does not stop them counting.
 */
export const RenewalSteps: Paragraph = () => {
  const [state] = useState((data: SampleData) => ({
    renewsOn: data.renewsOn,
    centre: data.centreName,
  }));

  return (
    <>
      <Paragraph id="renewal-steps-intro">To renew before {state.renewsOn}:</Paragraph>
      <Paragraph id="renewal-step-check" list="number">Check the plan and the price above.</Paragraph>
      <Paragraph id="renewal-step-pay" list="number">
        Pay online, or at the desk at {state.centre}.
      </Paragraph>
      <Paragraph id="renewal-step-pay-note" list="bullet" listLevel={1}>
        Card and direct debit are both accepted.
      </Paragraph>
      <Paragraph id="renewal-step-card" list="number">Collect your new card on your next visit.</Paragraph>
    </>
  );
};
