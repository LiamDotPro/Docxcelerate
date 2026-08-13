import { paragraph } from "docxcelerate";
import type { OfferData } from "../types.ts";

export const Greeting = paragraph<OfferData>({
  id: "greeting",
  render: (data) => `Dear ${data.applicantName},`,
});
