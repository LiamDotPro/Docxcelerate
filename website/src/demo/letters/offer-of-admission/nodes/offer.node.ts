import { paragraph } from "docxcelerate";
import type { OfferData } from "../types.ts";

export const Offer = paragraph<OfferData>({
  id: "offer",
  render: (data) =>
    `I am delighted to offer you a place on the ${data.programme} at ` +
    `${data.college}, beginning ${data.startDate}. Your application reference ` +
    `is ${data.offerRef}; please quote it in any correspondence with us.`,
});
