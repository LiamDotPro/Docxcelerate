/**
 * One dataset behind every node example on the site, so a reader moving from
 * paragraph to graph to section sees one letter taking shape rather than a
 * fresh cast per page.
 */
export interface SampleData {
  memberName: string;
  membershipRef: string;
  plan: string;
  centreName: string;
  renewsOn: string;
  /** Annual price, in pounds. */
  lastPrice: number;
  newPrice: number;
  /** Outstanding balance in pounds; 0 when the account is clear. */
  balance: number;
  managerName: string;
  signatureUrl: string;
  visitsByMonth: { month: string; visits: number }[];
  classMix: { label: string; share: number }[];
}

export const sampleData: SampleData = {
  memberName: "Adaeze Nkemelu",
  membershipRef: "RIV-88214",
  plan: "Peak Anytime",
  centreName: "Riverside Leisure Centre",
  renewsOn: "1 October 2026",
  lastPrice: 468,
  newPrice: 492,
  balance: 0,
  managerName: "Tomas Lindqvist",
  signatureUrl: "assets/signature-lindqvist.png",
  visitsByMonth: [
    { month: "Apr", visits: 11 },
    { month: "May", visits: 14 },
    { month: "Jun", visits: 9 },
    { month: "Jul", visits: 16 },
    { month: "Aug", visits: 18 },
    { month: "Sep", visits: 12 },
  ],
  classMix: [
    { label: "Swim", share: 42 },
    { label: "Strength", share: 33 },
    { label: "Classes", share: 25 },
  ],
};

/** Pounds, formatted the way the letter prints them. */
export function money(value: number): string {
  return value.toLocaleString("en-GB", {
    style: "currency",
    currency: "GBP",
    maximumFractionDigits: value % 1 === 0 ? 0 : 2,
  });
}
