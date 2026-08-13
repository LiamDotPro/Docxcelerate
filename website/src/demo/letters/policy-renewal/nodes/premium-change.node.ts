import { paragraph } from "docxcelerate";
import type { PolicyData } from "../types.ts";

/**
 * A price rise is the sentence customers actually read. Computing the
 * direction and the percentage locally means it can never contradict the
 * figures printed beside it.
 */
export const PremiumChange = paragraph<PolicyData>({
  id: "premium-change",
  render: (data) => {
    const delta = data.newPremium - data.lastPremium;
    const percent = Math.abs((delta / data.lastPremium) * 100).toFixed(1);

    if (Math.abs(delta) < 0.01) {
      return `Your premium is unchanged at ${money(data.newPremium)} a year.`;
    }

    const direction = delta > 0 ? "risen" : "fallen";
    return `Your premium has ${direction} by ${percent}%, from ` +
      `${money(data.lastPremium)} to ${money(data.newPremium)} a year. ` +
      `That works out at ${money(data.newPremium / 12)} a month.`;
  },
});

function money(value: number): string {
  return value.toLocaleString("en-GB", { style: "currency", currency: "GBP" });
}
