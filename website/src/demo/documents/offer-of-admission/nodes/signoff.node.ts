import { paragraph } from "docxcelerate";
import type { OfferData } from "../types.ts";

export const SignOff = paragraph<OfferData>({
  id: "sign-off",
  render: (data) =>
    `Yours sincerely, ${data.signatory.name}, ${data.signatory.title}, ${data.college}.`,
});
