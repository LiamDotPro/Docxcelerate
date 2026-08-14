import { paragraph } from "docxcelerate";
import type { PolicyData } from "../types.ts";

export const Renewal = paragraph<PolicyData>({
  id: "renewal",
  render: (data) =>
    `Your ${data.cover.toLowerCase()} policy ${data.policyNumber} renews on ` +
    `${data.renewalDate}. Unless you tell us otherwise, cover continues ` +
    `automatically with an excess of ${money(data.excess)}.`,
});

function money(value: number): string {
  return value.toLocaleString("en-GB", { style: "currency", currency: "GBP" });
}
