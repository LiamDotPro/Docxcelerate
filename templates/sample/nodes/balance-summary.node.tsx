import { dataRef, derive, Paragraph, useState } from "docxcelerate/document";
import type { DocumentData } from "../types.ts";

/**
 * A figure formatted per recipient rather than per build.
 *
 * `useFormat` would format it here, which is right whenever the value is known
 * now. This one is not: a published document is written for people whose
 * balances nobody has looked up yet. So the formatting is a deriver the engine
 * runs per document, and the text refers to what it produced.
 */
export const BalanceSummary: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    city: data.city,
  }));

  return (
    <Paragraph
      id="balance-summary"
      derivers={[
        derive("currencyLabel", {
          output: "balanceDueLabel",
          inputs: [dataRef("balanceDue")],
        }),
      ]}
    >
      Your current balance for {state.city} is {"{{derived.balanceDueLabel}}"}.
    </Paragraph>
  );
};
