import { paragraph } from "docxcelerate";
import type { PolicyData } from "../types.ts";

export const Greeting = paragraph<PolicyData>({
  id: "greeting",
  render: (data) => `Dear ${data.holderName},`,
});
