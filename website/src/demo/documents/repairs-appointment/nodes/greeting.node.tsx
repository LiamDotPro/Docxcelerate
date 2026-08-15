import { Paragraph, useState } from "docxcelerate/template";
import type { RepairsData } from "../types.ts";

export const Greeting: Paragraph = () => {
  const [state] = useState((data: RepairsData) => ({ name: data.residentName }));

  return <Paragraph id="greeting">Dear {state.name},</Paragraph>;
};
