/** @jsxImportSource docxcelerate/template */
import { Paragraph, useState } from "docxcelerate/template";
import type { OfferData } from "../types.ts";

export const Greeting: Paragraph = () => {
  const [state] = useState((data: OfferData) => ({ name: data.applicantName }));

  return <Paragraph id="greeting">Dear {state.name},</Paragraph>;
};
