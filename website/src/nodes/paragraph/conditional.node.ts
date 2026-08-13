import { paragraph } from "docxcelerate";
import { money, type SampleData } from "../sample-data.ts";

/**
 * Branching lives in the render, not the template: one node and one id,
 * whichever branch a member falls down.
 */
export const PriceChange = paragraph<SampleData>({
  id: "price-change",
  render: (data) => {
    const delta = data.newPrice - data.lastPrice;

    if (delta === 0) {
      return `Your ${data.plan} membership renews at ${money(data.newPrice)} ` +
        `a year — the same price you paid last year.`;
    }

    const direction = delta > 0 ? "rising" : "falling";
    const percent = Math.abs((delta / data.lastPrice) * 100).toFixed(1);

    return `Your ${data.plan} membership is ${direction} by ${percent}%, from ` +
      `${money(data.lastPrice)} to ${money(data.newPrice)} a year. That is ` +
      `${money(data.newPrice / 12)} a month from ${data.renewsOn}.`;
  },
});
