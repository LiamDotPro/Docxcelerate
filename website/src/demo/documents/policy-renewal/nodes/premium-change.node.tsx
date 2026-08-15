import { Paragraph, useFormat, useState } from "docxcelerate/template";
import type { PolicyData } from "../types.ts";

/**
 * A price rise is the sentence customers actually read. Working out the
 * direction and the percentage in the state initializer means the wording can
 * never contradict the figures printed beside it, because both come from the
 * same computation.
 */
export const PremiumChange: Paragraph = () => {
  const { currency } = useFormat("en-GB");
  const [state] = useState((data: PolicyData) => {
    const delta = data.newPremium - data.lastPremium;

    return {
      delta,
      lastPremium: data.lastPremium,
      newPremium: data.newPremium,
      percent: Math.abs((delta / data.lastPremium) * 100).toFixed(1),
    };
  });

  if (Math.abs(state.delta) < 0.01) {
    return (
      <Paragraph id="premium-held">
        Your premium is unchanged at {currency(state.newPremium)} a year.
      </Paragraph>
    );
  }

  return (
    <Paragraph id="premium-change">
      Your premium has {state.delta > 0 ? "risen" : "fallen"} by {state.percent}%, from{" "}
      {currency(state.lastPremium)} to {currency(state.newPremium)} a year. That works out
      at {currency(state.newPremium / 12)} a month.
    </Paragraph>
  );
};
