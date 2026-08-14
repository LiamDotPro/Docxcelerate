/** @jsxImportSource docxcelerate/template */
import { Paragraph, useFormat, useState } from "docxcelerate/template";
import type { PolicyData } from "../types.ts";

export const Renewal: Paragraph = () => {
  const { currency } = useFormat("en-GB");
  const [state] = useState((data: PolicyData) => ({
    cover: data.cover.toLowerCase(),
    policyNumber: data.policyNumber,
    renewalDate: data.renewalDate,
    excess: data.excess,
  }));

  return (
    <Paragraph id="renewal-terms">
      Your {state.cover} policy {state.policyNumber} renews on {state.renewalDate}. Unless
      you tell us otherwise, cover continues automatically with an excess of{" "}
      {currency(state.excess)}.
    </Paragraph>
  );
};
