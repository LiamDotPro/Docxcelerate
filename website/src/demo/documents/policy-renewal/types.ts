export interface PolicyData {
  holderName: string;
  policyNumber: string;
  cover: "Contents" | "Buildings" | "Contents and buildings";
  renewalDate: string;
  lastPremium: number;
  newPremium: number;
  excess: number;
}
