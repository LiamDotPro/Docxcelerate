/** @jsxImportSource docxcelerate/template */
import { Paragraph, useState } from "docxcelerate/template";
import { money, type SampleData } from "../sample-data.ts";

/**
 * Branching is an ordinary `if`. The component decides what it is before it
 * says anything, and each arm returns the node that arm means — one id per
 * outcome, so a reader of the resolved document can tell which one they got.
 */
export const PriceChange: Paragraph = () => {
  const [state] = useState((data: SampleData) => {
    const delta = data.newPrice - data.lastPrice;

    return {
      plan: data.plan,
      renewsOn: data.renewsOn,
      lastPrice: data.lastPrice,
      newPrice: data.newPrice,
      delta,
      percent: Math.abs((delta / data.lastPrice) * 100).toFixed(1),
    };
  });

  if (state.delta === 0) {
    return (
      <Paragraph id="price-held">
        Your {state.plan} membership renews at {money(state.newPrice)} a year — the same
        price you paid last year.
      </Paragraph>
    );
  }

  return (
    <Paragraph id="price-change">
      Your {state.plan} membership is {state.delta > 0 ? "rising" : "falling"} by{" "}
      {state.percent}%, from {money(state.lastPrice)} to {money(state.newPrice)} a year.
      That is {money(state.newPrice / 12)} a month from {state.renewsOn}.
    </Paragraph>
  );
};
