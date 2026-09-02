import type { DeriverDefinitions } from "docxcelerate/document";

export const derivers = {
  currencyLabel: ([amount]) =>
    new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
    }).format(Number(amount ?? 0)),
} satisfies DeriverDefinitions;

export default derivers;
